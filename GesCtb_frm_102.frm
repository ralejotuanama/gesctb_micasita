VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_Period_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4395
   ClientLeft      =   4245
   ClientTop       =   4380
   ClientWidth     =   6405
   Icon            =   "GesCtb_frm_102.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4365
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6405
      _Version        =   65536
      _ExtentX        =   11298
      _ExtentY        =   7699
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   1125
         Left            =   30
         TabIndex        =   8
         Top             =   3180
         Width           =   6315
         _Version        =   65536
         _ExtentX        =   11139
         _ExtentY        =   1984
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
         Begin VB.ComboBox cmb_BlqReg 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   720
            Width           =   4725
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1530
            TabIndex        =   2
            Top             =   60
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1530
            TabIndex        =   3
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
         Begin VB.Label Label7 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   60
            TabIndex        =   20
            Top             =   390
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Inicio:"
            Height          =   285
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label26 
            Caption         =   "Bloqueo Registros:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   720
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   6315
         _Version        =   65536
         _ExtentX        =   11139
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
            Left            =   5670
            Picture         =   "GesCtb_frm_102.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_102.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   795
         Left            =   30
         TabIndex        =   11
         Top             =   2340
         Width           =   6315
         _Version        =   65536
         _ExtentX        =   11139
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   4725
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1530
            TabIndex        =   1
            Top             =   420
            Width           =   1125
            _Version        =   196608
            _ExtentX        =   1984
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
            MinValue        =   "2009"
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
            Left            =   60
            TabIndex        =   13
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Mes:"
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   6315
         _Version        =   65536
         _ExtentX        =   11139
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
            Height          =   480
            Left            =   630
            TabIndex        =   15
            Top             =   90
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de Períodos"
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
            Left            =   90
            Picture         =   "GesCtb_frm_102.frx":0890
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   855
         Left            =   30
         TabIndex        =   16
         Top             =   1440
         Width           =   6315
         _Version        =   65536
         _ExtentX        =   11139
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
         Begin Threed.SSPanel pnl_NomEmp 
            Height          =   345
            Left            =   1530
            TabIndex        =   17
            Top             =   60
            Width           =   4725
            _Version        =   65536
            _ExtentX        =   8334
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "SSPanel10"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_TipPer 
            Height          =   345
            Left            =   1530
            TabIndex        =   21
            Top             =   450
            Width           =   4725
            _Version        =   65536
            _ExtentX        =   8334
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "SSPanel10"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo Período:"
            Height          =   285
            Left            =   60
            TabIndex        =   22
            Top             =   450
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "Empresa:"
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_Period_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_SitOpe     As Integer
Dim l_int_SitCtb     As Integer

Private Sub cmb_PerMes_Click()
   Call gs_SetFocus(ipp_PerAno)
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PerMes_Click
   End If
End Sub

Private Sub cmb_BlqReg_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_BlqReg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BlqReg_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el mes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   If moddat_g_str_CodMod = 2 Then
      If Month(CDate(ipp_FecIni.Text)) <> Month(CDate(ipp_FecFin.Text)) Then
         MsgBox "El rango de fechas debe comprender el mismo mes.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIni)
         Exit Sub
      End If
   End If
   
   If cmb_BlqReg.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Bloqueo de Registros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_BlqReg)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = "SELECT * FROM CTB_PERMES WHERE "
      g_str_Parame = g_str_Parame & "PERMES_CODEMP  = '" & moddat_g_str_CodGrp & "' AND "
      g_str_Parame = g_str_Parame & "PERMES_TIPPER = " & moddat_g_str_CodMod & " AND "
      g_str_Parame = g_str_Parame & "PERMES_CODANO = " & ipp_PerAno.Text & " AND "
      g_str_Parame = g_str_Parame & "PERMES_CODMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " "
                     
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing

         MsgBox "El Período ya ha sido registrado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PerMes)
         Exit Sub
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      'Validar el Rango de Fechas
      g_str_Parame = "SELECT * FROM CTB_PERMES WHERE "
      g_str_Parame = g_str_Parame & "PERMES_CODEMP  = '" & moddat_g_str_CodGrp & "' AND "
      g_str_Parame = g_str_Parame & "PERMES_TIPPER = " & moddat_g_str_CodMod & " AND "
      g_str_Parame = g_str_Parame & "( (PERMES_FECINI <= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND PERMES_FECFIN >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND PERMES_FECINI <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND PERMES_FECFIN >= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ") OR  "
      g_str_Parame = g_str_Parame & "(PERMES_FECINI >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND PERMES_FECINI <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND PERMES_FECFIN >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND PERMES_FECFIN <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ") OR  "
      g_str_Parame = g_str_Parame & "(PERMES_FECINI >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND PERMES_FECINI <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND PERMES_FECFIN >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND PERMES_FECFIN >= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ") OR  "
      g_str_Parame = g_str_Parame & "(PERMES_FECINI <= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND PERMES_FECINI <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND PERMES_FECFIN >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND PERMES_FECFIN <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ") ) "
      g_str_Parame = g_str_Parame & "ORDER BY PERMES_FECINI ASC "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing

         MsgBox "El Rango de Fechas ya ha sido registrado en otro Período.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIni)
         Exit Sub
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      'Validar que no exista otro Período Vigente
      g_str_Parame = "SELECT * FROM CTB_PERMES WHERE "
      g_str_Parame = g_str_Parame & "PERMES_CODEMP  = '" & moddat_g_str_CodGrp & "' AND "
      g_str_Parame = g_str_Parame & "PERMES_TIPPER = " & moddat_g_str_CodMod & " AND "
      g_str_Parame = g_str_Parame & "PERMES_SITUAC = 1 "
                     
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing

         MsgBox "Existe otro Período en Proceso.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PerMes)
         Exit Sub
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CTB_PERMES ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
      g_str_Parame = g_str_Parame & moddat_g_str_CodMod & ", "
      g_str_Parame = g_str_Parame & ipp_PerAno.Text & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & CStr(cmb_BlqReg.ItemData(cmb_BlqReg.ListIndex)) & ", "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ")"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CTB_PERMES. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   moddat_g_int_FlgAct = 2
   
   Unload Me
   
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   
   pnl_NomEmp.Caption = moddat_g_str_CodGrp & " - " & moddat_g_str_DesGrp
   pnl_TipPer.Caption = moddat_g_str_CodMod & " - " & moddat_g_str_DesMod
   
   Call fs_Inicio
   Call fs_Limpia
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = "SELECT * FROM CTB_PERMES WHERE "
      g_str_Parame = g_str_Parame & "PERMES_CODEMP = '" & moddat_g_str_CodGrp & "' AND "
      g_str_Parame = g_str_Parame & "PERMES_TIPPER = " & moddat_g_str_CodMod & " AND "
      g_str_Parame = g_str_Parame & "PERMES_CODANO = " & moddat_g_str_Codigo & " AND "
      g_str_Parame = g_str_Parame & "PERMES_CODMES = " & moddat_g_str_CodIte & "  "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
      
         Call gs_BuscarCombo_Item(cmb_PerMes, g_rst_Princi!PERMES_CODMES)
         ipp_PerAno.Text = CStr(g_rst_Princi!PERMES_CODANO)
         
         cmb_PerMes.Enabled = False
         ipp_PerAno.Enabled = False
         
         ipp_FecIni.Text = gf_FormatoFecha(CStr(g_rst_Princi!PERMES_FECINI))
         ipp_FecFin.Text = gf_FormatoFecha(CStr(g_rst_Princi!PERMES_FECFIN))
         
         ipp_FecIni.Enabled = False
         ipp_FecFin.Enabled = False
         
         Call gs_BuscarCombo_Item(cmb_BlqReg, g_rst_Princi!PERMES_BLQREG)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   Call moddat_gs_Carga_LisIte_Combo(cmb_BlqReg, 1, "214")
End Sub

Private Sub fs_Limpia()
   Call gs_BuscarCombo_Item(cmb_PerMes, Month(date))
   ipp_PerAno.Text = Format(Year(date), "0000")
   
   ipp_FecIni.Text = "01/" & Format(Month(date), "00") & "/" & Format(Year(date), "0000")
   ipp_FecFin.Text = Format(ff_Ultimo_Dia_Mes(Month(date), Year(date)), "00") & "/" & Format(Month(date), "00") & "/" & Format(Year(date), "0000")
   
   cmb_BlqReg.ListIndex = -1
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_BlqReg)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

