VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_RecPrv_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   Icon            =   "GesCtb_frm_905.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3225
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5865
      _Version        =   65536
      _ExtentX        =   10345
      _ExtentY        =   5689
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
         TabIndex        =   6
         Top             =   60
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
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
            Left            =   600
            TabIndex        =   7
            Top             =   120
            Width           =   2865
            _Version        =   65536
            _ExtentX        =   5054
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Recalculo de Provisiones"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Picture         =   "GesCtb_frm_905.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   780
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
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
            Left            =   5160
            Picture         =   "GesCtb_frm_905.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_905.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Procesar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1215
         Left            =   30
         TabIndex        =   9
         Top             =   1470
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
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
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   3975
         End
         Begin VB.ComboBox cmb_Period 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   450
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   810
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
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   90
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   450
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   840
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   13
         Top             =   2730
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
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
            TabIndex        =   14
            Top             =   60
            Width           =   5655
            _Version        =   65536
            _ExtentX        =   9975
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SSPanel2"
            ForeColor       =   16777215
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
            FloodType       =   1
            FloodColor      =   49152
            Font3D          =   2
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_RecPrv_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_str_PerAno         As String
Dim l_str_PerMes         As String
Dim l_str_CodEmp         As String

Private Sub cmd_Proces_Click()
Dim r_str_Periodo      As String
Dim r_str_FecIni       As String
Dim r_str_FecFin       As String
   
   If cmb_Empres.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Empres)
      Exit Sub
   End If
   If cmb_Period.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Period)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe registrar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   l_str_PerAno = Format(ipp_PerAno.Text, "0000")
   l_str_PerMes = Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00")
   l_str_CodEmp = l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo
   r_str_Periodo = moddat_gf_ConsultaPerMesActivo(l_str_CodEmp, 2, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   
   If modctb_int_PerMes = 1 Then
      l_str_PerMes = 12
      l_str_PerAno = modctb_int_PerAno - 1
   Else
      l_str_PerMes = modctb_int_PerMes - 1
   End If
   
   If CInt(l_str_PerAno) <> CInt(ipp_PerAno.Text) Then
      MsgBox "Año registrado no puede reprocesarse.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   If l_str_PerMes <> cmb_Period.ItemData(cmb_Period.ListIndex) Then
      MsgBox "Mes seleccionado no puede reprocesarse.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Period)
      Exit Sub
   End If
   modctb_str_FecFin = fp_UltimoDia(l_str_PerMes, l_str_PerAno)
   r_str_FecIni = "01" & "/" & Format(modctb_int_PerMes, "00") & "/" & Format(modctb_int_PerAno, "0000")
   r_str_FecFin = "20" & "/" & Format(modctb_int_PerMes, "00") & "/" & Format(modctb_int_PerAno, "0000")
   
   If CDate(moddat_g_str_FecSis) < CDate(r_str_FecIni) Then
      MsgBox "Periodo seleccionado ya no puede reprocesarse.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Period)
      Exit Sub
   End If
   If CDate(moddat_g_str_FecSis) > CDate(r_str_FecFin) Then
      MsgBox "Periodo seleccionado ya no puede reprocesarse.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Period)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de Realizar el Re-Proceso de Provisiones?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
    
   cmd_Proces.Enabled = False
   Screen.MousePointer = 11
   pnl_BarPro.FloodPercent = 0
   
   'Cálculo de Provisiones
   Call modprc_ctbp1009(l_str_CodEmp, l_str_PerMes, CInt(l_str_PerAno), modctb_str_FecFin, pnl_BarPro)
   pnl_BarPro.FloodPercent = 100
    
   MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   Screen.MousePointer = 0
   cmd_Proces.Enabled = True
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_SetFocus(cmb_Empres)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_Period, 1, "033")
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   ipp_PerAno = Mid(date, 7, 4)
   cmb_Empres.ListIndex = 0
   pnl_BarPro.FloodPercent = 0
End Sub

'*** Funcion para Devolver el Ultimo Dia de cualquier Mes
Private Function fp_UltimoDia(ByVal p_Fec_xMes As String, p_Fec_xAno As String) As String
Dim r_Int_Dia As Integer
   
   Select Case CInt(p_Fec_xMes)
       Case 1, 3, 5, 7, 8, 10, 12
            r_Int_Dia = 31
       Case 2
            If (CInt(p_Fec_xAno) - 2008) Mod 4 = 0 Then
                'Año bisiesto 2008, 2012,2016, 2020...etc
                    r_Int_Dia = 29
                Else
                    r_Int_Dia = 28
            End If
       Case 4, 6, 9, 11
            r_Int_Dia = 30
   End Select
   
   fp_UltimoDia = Str(r_Int_Dia) + "/" + p_Fec_xMes + "/" + p_Fec_xAno
End Function

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Period)
   End If
End Sub

Private Sub cmb_Period_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub

