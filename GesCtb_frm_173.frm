VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mat_Produc_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7275
   ClientLeft      =   4230
   ClientTop       =   2730
   ClientWidth     =   9300
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7245
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9285
      _Version        =   65536
      _ExtentX        =   16378
      _ExtentY        =   12779
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
         TabIndex        =   16
         Top             =   60
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
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
            TabIndex        =   17
            Top             =   60
            Width           =   5955
            _Version        =   65536
            _ExtentX        =   10504
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Cuentas Contables por Producto"
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
            Picture         =   "GesCtb_frm_173.frx":0000
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   18
         Top             =   780
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   60
            Picture         =   "GesCtb_frm_173.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   8580
            Picture         =   "GesCtb_frm_173.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   2145
         Left            =   30
         TabIndex        =   19
         Top             =   3420
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
         _ExtentY        =   3784
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
         Begin VB.CheckBox chk_ClaGar 
            Caption         =   "No Aplica"
            Height          =   315
            Left            =   1770
            TabIndex        =   8
            Top             =   1800
            Width           =   3525
         End
         Begin VB.ComboBox cmb_ClaGar 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1440
            Width           =   7365
         End
         Begin VB.CheckBox chk_SitCre 
            Caption         =   "No Aplica"
            Height          =   315
            Left            =   1770
            TabIndex        =   6
            Top             =   1110
            Width           =   4995
         End
         Begin VB.ComboBox cmb_SitCre 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   750
            Width           =   7365
         End
         Begin VB.CheckBox chk_TipCre 
            Caption         =   "No Aplica"
            Height          =   315
            Left            =   1770
            TabIndex        =   4
            Top             =   420
            Width           =   4155
         End
         Begin VB.ComboBox cmb_TipCre 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   60
            Width           =   7365
         End
         Begin VB.Label Label5 
            Caption         =   "Clase de Garantía:"
            Height          =   285
            Left            =   60
            TabIndex        =   22
            Top             =   1440
            Width           =   1515
         End
         Begin VB.Label Label3 
            Caption         =   "Situación de Crédito:"
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   750
            Width           =   1515
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Crédito:"
            Height          =   285
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1095
         Left            =   30
         TabIndex        =   23
         Top             =   2280
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
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
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   7365
         End
         Begin VB.ComboBox cmb_ConCtb 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   7365
         End
         Begin VB.ComboBox cmb_CtaCtb 
            Height          =   315
            Left            =   1770
            TabIndex        =   2
            Text            =   "cmb_CtaCtb"
            Top             =   720
            Width           =   7365
         End
         Begin VB.Label Label7 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1605
         End
         Begin VB.Label Label6 
            Caption         =   "Concepto Contable:"
            Height          =   255
            Left            =   60
            TabIndex        =   25
            Top             =   390
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Contable:"
            Height          =   255
            Left            =   60
            TabIndex        =   24
            Top             =   720
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   27
         Top             =   1470
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1770
            TabIndex        =   28
            Top             =   60
            Width           =   7365
            _Version        =   65536
            _ExtentX        =   12991
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_EmpGrp 
            Height          =   315
            Left            =   1770
            TabIndex        =   29
            Top             =   390
            Width           =   7365
            _Version        =   65536
            _ExtentX        =   12991
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
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
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label8 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   31
            Top             =   60
            Width           =   1605
         End
         Begin VB.Label Label4 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   30
            Top             =   390
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   765
         Left            =   30
         TabIndex        =   32
         Top             =   5610
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
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
         Begin VB.ComboBox cmb_EmpSeg 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   60
            Width           =   7365
         End
         Begin VB.CheckBox chk_EmpSeg 
            Caption         =   "No Aplica"
            Height          =   315
            Left            =   1770
            TabIndex        =   10
            Top             =   420
            Width           =   4155
         End
         Begin VB.Label Label11 
            Caption         =   "Empresa Seguro:"
            Height          =   285
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   765
         Left            =   30
         TabIndex        =   34
         Top             =   6420
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
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
         Begin VB.CheckBox chk_GasCie 
            Caption         =   "No Aplica"
            Height          =   315
            Left            =   1770
            TabIndex        =   12
            Top             =   420
            Width           =   4155
         End
         Begin VB.ComboBox cmb_GasCie 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   60
            Width           =   7365
         End
         Begin VB.Label Label9 
            Caption         =   "Gasto de Cierre:"
            Height          =   285
            Left            =   60
            TabIndex        =   35
            Top             =   60
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_Mat_Produc_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_TopNiv     As Integer
Dim l_int_FlgCmb     As Integer
Dim l_str_CtaCtb     As String
Dim l_arr_ParEmp()   As moddat_tpo_Genera
Dim l_arr_ConCtb()   As moddat_tpo_Genera
Dim l_arr_CtaCtb()   As moddat_tpo_Genera
Dim l_arr_TipCre()   As moddat_tpo_Genera
Dim l_arr_SitCre()   As moddat_tpo_Genera
Dim l_arr_ClaGar()   As moddat_tpo_Genera
Dim l_arr_EmpSeg()   As moddat_tpo_Genera

Private Sub chk_ClaGar_Click()
   If chk_ClaGar.Value = 1 Then
      cmb_ClaGar.ListIndex = -1
      cmb_ClaGar.Enabled = False
   Else
      cmb_ClaGar.Enabled = True
   End If
End Sub

Private Sub chk_EmpSeg_Click()
   If chk_EmpSeg.Value = 1 Then
      cmb_EmpSeg.ListIndex = -1
      cmb_EmpSeg.Enabled = False
   Else
      cmb_EmpSeg.Enabled = True
   End If
End Sub

Private Sub chk_GasCie_Click()
   If chk_GasCie.Value = 1 Then
      cmb_GasCie.ListIndex = -1
      cmb_GasCie.Enabled = False
   Else
      cmb_GasCie.Enabled = True
   End If
End Sub

Private Sub chk_SitCre_Click()
   If chk_SitCre.Value = 1 Then
      cmb_SitCre.ListIndex = -1
      cmb_SitCre.Enabled = False
   Else
      cmb_SitCre.Enabled = True
   End If
End Sub

Private Sub chk_TipCre_Click()
   If chk_TipCre.Value = 1 Then
      cmb_TipCre.ListIndex = -1
      cmb_TipCre.Enabled = False
   Else
      cmb_TipCre.Enabled = True
   End If
End Sub

Private Sub cmb_ClaGar_Click()
   If cmb_EmpSeg.Enabled Then
      Call gs_SetFocus(cmb_EmpSeg)
   ElseIf cmb_GasCie.Enabled Then
      Call gs_SetFocus(cmb_GasCie)
   Else
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_ClaGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ClaGar_Click
   End If
End Sub

Private Sub cmb_ConCtb_Click()
   Call gs_SetFocus(cmb_CtaCtb)
End Sub

Private Sub cmb_ConCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ConCtb_Click
   End If
End Sub

Private Sub cmb_CtaCtb_Change()
   l_str_CtaCtb = cmb_CtaCtb.Text
   
   cmb_CtaCtb.SelLength = Len(l_str_CtaCtb)
End Sub

Private Sub cmb_CtaCtb_Click()
   If cmb_CtaCtb.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_TipCre)
      End If
   End If
End Sub

Private Sub cmb_CtaCtb_GotFocus()
   Call SendMessage(cmb_CtaCtb.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   
   l_int_FlgCmb = True
End Sub

Private Sub cmb_CtaCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_CtaCtb, l_str_CtaCtb)
      l_int_FlgCmb = True
      
      If cmb_CtaCtb.ListIndex > -1 Then
         l_str_CtaCtb = ""
      End If
      
      Call gs_SetFocus(cmb_TipCre)
   End If
End Sub

Private Sub cmb_CtaCtb_LostFocus()
   Call SendMessage(cmb_CtaCtb.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_EmpSeg_Click()
   If cmb_GasCie.Enabled Then
      Call gs_SetFocus(cmb_GasCie)
   Else
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_EmpSeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpSeg_Click
   End If
End Sub

Private Sub cmb_GasCie_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_GasCie_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_GasCie_Click
   End If
End Sub

Private Sub cmb_SitCre_Click()
   If cmb_ClaGar.Enabled Then
      Call gs_SetFocus(cmb_ClaGar)
   ElseIf cmb_EmpSeg.Enabled Then
      Call gs_SetFocus(cmb_EmpSeg)
   ElseIf cmb_GasCie.Enabled Then
      Call gs_SetFocus(cmb_GasCie)
   Else
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_SitCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SitCre_Click
   End If
End Sub

Private Sub cmb_TipCre_Click()
   If cmb_SitCre.Enabled Then
      Call gs_SetFocus(cmb_SitCre)
   ElseIf cmb_ClaGar.Enabled Then
      Call gs_SetFocus(cmb_ClaGar)
   ElseIf cmb_EmpSeg.Enabled Then
      Call gs_SetFocus(cmb_EmpSeg)
   ElseIf cmb_GasCie.Enabled Then
      Call gs_SetFocus(cmb_GasCie)
   Else
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_TipCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipCre_Click
   End If
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(cmb_ConCtb)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_TipCre     As String
   Dim r_str_SitCre     As String
   Dim r_str_ClaGar     As String
   Dim r_str_EmpSeg     As String
   Dim r_str_GasCie     As String

   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   If cmb_ConCtb.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Concepto Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ConCtb)
      Exit Sub
   End If
   
   If cmb_CtaCtb.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Cuenta Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaCtb)
      Exit Sub
   End If
   
   If chk_TipCre.Value = 0 Then
      If cmb_TipCre.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Crédito.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipCre)
         Exit Sub
      End If
      
      r_str_TipCre = l_arr_TipCre(cmb_TipCre.ListIndex + 1).Genera_Codigo
   Else
      r_str_TipCre = "999"
   End If

   If chk_SitCre.Value = 0 Then
      If cmb_SitCre.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Situación de Crédito.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SitCre)
         Exit Sub
      End If
      
      r_str_SitCre = l_arr_SitCre(cmb_SitCre.ListIndex + 1).Genera_Codigo
   Else
      r_str_SitCre = "999"
   End If
   
   If chk_ClaGar.Value = 0 Then
      If cmb_ClaGar.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Clase de Garantía.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ClaGar)
         Exit Sub
      End If
      
      r_str_ClaGar = l_arr_ClaGar(cmb_ClaGar.ListIndex + 1).Genera_Codigo
   Else
      r_str_ClaGar = "999"
   End If

   If chk_EmpSeg.Value = 0 Then
      If cmb_EmpSeg.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Empresa de Seguros.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_EmpSeg)
         Exit Sub
      End If
      
      r_str_EmpSeg = l_arr_EmpSeg(cmb_EmpSeg.ListIndex + 1).Genera_Codigo
   Else
      r_str_EmpSeg = "999999"
   End If

   If chk_GasCie.Value = 0 Then
      If cmb_GasCie.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Gasto de Cierre.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_GasCie)
         Exit Sub
      End If
      
      r_str_GasCie = Format(cmb_GasCie.ItemData(cmb_GasCie.ListIndex), "00")
   Else
      r_str_GasCie = "99"
   End If
   
   
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAPRD WHERE "
      g_str_Parame = g_str_Parame & "CTAPRD_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_TIPCRE = '" & r_str_TipCre & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_SITCRE = '" & r_str_SitCre & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_CLAGAR = '" & r_str_ClaGar & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_EMPSEG = '" & r_str_EmpSeg & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_GASCIE = " & r_str_GasCie & " AND "
      g_str_Parame = g_str_Parame & "CTAPRD_TIPMON = " & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & " AND "
      g_str_Parame = g_str_Parame & "CTAPRD_CONCTB = '" & l_arr_ConCtb(cmb_ConCtb.ListIndex + 1).Genera_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_EMPGRP = '" & moddat_g_str_CodEmp & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
        
         MsgBox "El Concepto Contable para esta Moneda ya ha sido registrado. Por favor verifique e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CTB_CTAPRD ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & l_arr_ConCtb(cmb_ConCtb.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_TipCre & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_SitCre & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_ClaGar & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_EmpSeg & "', "
      g_str_Parame = g_str_Parame & r_str_GasCie & ", "
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodEmp & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_CtaCtb(cmb_CtaCtb.ListIndex + 1).Genera_Codigo & "', "
         
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ")"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   Screen.MousePointer = 0
   
   moddat_g_int_FlgAct = 2
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt
   
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_EmpGrp.Caption = moddat_g_str_RazSoc
   
   Call fs_Inicio

   Call gs_CentraForm(Me)
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = "SELECT * FROM CTB_CTAPRD WHERE CTAPRD_CODPRD = '" & moddat_g_str_CodPrd & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_TIPMON = " & CStr(moddat_g_int_TipMon) & " AND "
      g_str_Parame = g_str_Parame & "CTAPRD_CONCTB = '" & moddat_g_str_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_EMPGRP = '" & moddat_g_str_CodEmp & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_TIPCRE = '" & moddat_g_str_TipCre & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_SITCRE = '" & moddat_g_str_SitCre & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_CLAGAR = '" & moddat_g_str_ClaGar & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_EMPSEG = '" & moddat_g_str_CodGrp & "' AND "
      g_str_Parame = g_str_Parame & "CTAPRD_GASCIE= " & moddat_g_str_CodIte & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!CTAPRD_TIPMON)
         
         cmb_ConCtb.ListIndex = gf_Busca_Arregl(l_arr_ConCtb, Trim(g_rst_Princi!CTAPRD_CONCTB)) - 1
         cmb_CtaCtb.ListIndex = gf_Busca_Arregl(l_arr_CtaCtb, Trim(g_rst_Princi!CTAPRD_CTACTB)) - 1
         
         If g_rst_Princi!CTAPRD_TIPCRE = "999" Then
            chk_TipCre.Value = 1
            cmb_TipCre.ListIndex = -1
         Else
            cmb_TipCre.ListIndex = gf_Busca_Arregl(l_arr_TipCre, Trim(g_rst_Princi!CTAPRD_TIPCRE)) - 1
         End If
         
         If g_rst_Princi!CTAPRD_SITCRE = "999" Then
            chk_SitCre.Value = 1
            cmb_SitCre.ListIndex = -1
         Else
            cmb_SitCre.ListIndex = gf_Busca_Arregl(l_arr_SitCre, Trim(g_rst_Princi!CTAPRD_SITCRE)) - 1
         End If
         
         If g_rst_Princi!CTAPRD_CLAGAR = "999" Then
            chk_ClaGar.Value = 1
            cmb_ClaGar.ListIndex = -1
         Else
            cmb_ClaGar.ListIndex = gf_Busca_Arregl(l_arr_ClaGar, Trim(g_rst_Princi!CTAPRD_CLAGAR)) - 1
         End If
         
         If g_rst_Princi!CTAPRD_EMPSEG = "999999" Then
            chk_EmpSeg.Value = 1
            cmb_EmpSeg.ListIndex = -1
         Else
            cmb_EmpSeg.ListIndex = gf_Busca_Arregl(l_arr_EmpSeg, Trim(g_rst_Princi!CTAPRD_EMPSEG)) - 1
         End If
         
         If g_rst_Princi!CTAPRD_GASCIE = 99 Then
            chk_GasCie.Value = 1
            cmb_GasCie.ListIndex = -1
         Else
            Call gs_BuscarCombo_Item(cmb_GasCie, g_rst_Princi!CTAPRD_GASCIE)
         End If
         
         cmb_TipCre.Enabled = False
         chk_TipCre.Enabled = False
         cmb_SitCre.Enabled = False
         chk_SitCre.Enabled = False
         cmb_ClaGar.Enabled = False
         chk_ClaGar.Enabled = False
         cmb_EmpSeg.Enabled = False
         chk_EmpSeg.Enabled = False
         cmb_GasCie.Enabled = False
         chk_GasCie.Enabled = False
         cmb_TipMon.Enabled = False
         cmb_ConCtb.Enabled = False
         
         Call gs_SetFocus(cmb_CtaCtb)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   l_int_TopNiv = -1
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, moddat_g_str_CodEmp, "100", "001") Then
      l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
   End If
   
   'Call moddat_gs_Carga_CtaCtb(moddat_g_str_CodEmp, cmb_CtaCtb, l_arr_CtaCtb, 0, l_int_TopNiv, -1)
   
   Call modtac_gs_Carga_CtaCtb(moddat_g_str_CodEmp, cmb_CtaCtb, l_arr_CtaCtb, 0, l_int_TopNiv, -1)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
   Call moddat_gs_Carga_ConceptoCtb(l_arr_ConCtb, cmb_ConCtb, 2)
   
   Call moddat_gs_Carga_TipoCreditoCtb(l_arr_TipCre, cmb_TipCre)
   Call moddat_gs_Carga_SituacionCreditoCtb(l_arr_SitCre, cmb_SitCre, "4")
   Call moddat_gs_Carga_ClaGar(cmb_ClaGar, l_arr_ClaGar)
   
   Call moddat_gs_Carga_EmpSeg(cmb_EmpSeg, l_arr_EmpSeg)
   Call moddat_gs_Carga_LisIte_Combo(cmb_GasCie, 1, "265")
End Sub


