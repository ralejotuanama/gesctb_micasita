VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Pro_AsiDev_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   2760
   ClientLeft      =   5700
   ClientTop       =   3285
   ClientWidth     =   7935
   Icon            =   "GesCtb_frm_803.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2805
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7995
      _Version        =   65536
      _ExtentX        =   14102
      _ExtentY        =   4948
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
         Left            =   30
         TabIndex        =   4
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   480
            Left            =   630
            TabIndex        =   5
            Top             =   60
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Procesos - Asiento de Devengados (Cred. Hipot.)"
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
            Picture         =   "GesCtb_frm_803.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   6
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
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_803.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Procesar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7230
            Picture         =   "GesCtb_frm_803.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   7
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
            TabIndex        =   8
            Top             =   390
            Width           =   6285
            _Version        =   65536
            _ExtentX        =   11086
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin VB.Label Label3 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Per�odo:"
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   9
            Top             =   390
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   11
         Top             =   2280
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
            TabIndex        =   12
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
Attribute VB_Name = "frm_Pro_AsiDev_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Empres()      As moddat_tpo_Genera

Private Sub cmb_Empres_Click()
   If cmb_Empres.ListIndex > -1 Then
      Screen.MousePointer = 11
      pnl_Period.Caption = moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 2, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
      Call gs_SetFocus(cmd_Proces)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmd_Proces_Click()
Dim r_lng_TotErr     As Long

   If cmb_Empres.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Empres)
      Exit Sub
   End If
   If moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(modctb_str_FecFin), "yyyymmdd"), 2) = 0 Then
      MsgBox "Debe ingresar el Tipo de Cambio del Cierre.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("Est� seguro de ejecutar el proceso de Asientos Devengados?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   cmd_Proces.Enabled = False
   Screen.MousePointer = 11
   
   Call modprc_ctbp1012(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, modctb_int_PerMes, modctb_int_PerAno, modctb_str_FecFin)
   
   MsgBox "Proceso concluido.", vbInformation, modgen_g_str_NomPlt
   Screen.MousePointer = 0
   cmd_Proces.Enabled = True
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Empres)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
End Sub

Private Sub fs_Limpia()
   cmb_Empres.ListIndex = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

