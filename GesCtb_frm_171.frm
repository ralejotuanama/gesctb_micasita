VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mat_CtaBco_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3780
   ClientLeft      =   855
   ClientTop       =   2265
   ClientWidth     =   9300
   Icon            =   "GesCtb_frm_171.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3765
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9285
      _Version        =   65536
      _ExtentX        =   16378
      _ExtentY        =   6641
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   7
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
         Begin Threed.SSPanel pnl_CodBan 
            Height          =   315
            Left            =   1770
            TabIndex        =   8
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
            TabIndex        =   9
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
         Begin VB.Label Label4 
            Caption         =   "Banco:"
            Height          =   255
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   10
            Top             =   390
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   12
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
            TabIndex        =   13
            Top             =   60
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Cuentas Contables por Cuentas Bancarias"
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
            Picture         =   "GesCtb_frm_171.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   14
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   8580
            Picture         =   "GesCtb_frm_171.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_171.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1425
         Left            =   30
         TabIndex        =   15
         Top             =   2280
         Width           =   9195
         _Version        =   65536
         _ExtentX        =   16219
         _ExtentY        =   2514
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
         Begin VB.ComboBox cmb_CtaBan 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   7365
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   7365
         End
         Begin VB.ComboBox cmb_ConCtb 
            Height          =   315
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   7365
         End
         Begin VB.ComboBox cmb_CtaCtb 
            Height          =   315
            Left            =   1770
            TabIndex        =   3
            Text            =   "cmb_CtaCtb"
            Top             =   1050
            Width           =   7365
         End
         Begin VB.Label Label6 
            Caption         =   "Cuenta Bancaria:"
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   1605
         End
         Begin VB.Label Label2 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   390
            Width           =   1605
         End
         Begin VB.Label Label3 
            Caption         =   "Concepto Contable:"
            Height          =   255
            Left            =   60
            TabIndex        =   17
            Top             =   720
            Width           =   1605
         End
         Begin VB.Label Label5 
            Caption         =   "Cuenta Contable:"
            Height          =   255
            Left            =   60
            TabIndex        =   16
            Top             =   1050
            Width           =   1605
         End
      End
   End
End
Attribute VB_Name = "frm_Mat_CtaBco_02"
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
Dim l_arr_CtaBan()   As moddat_tpo_Genera

Private Sub cmb_ConCtb_Click()
   Call gs_SetFocus(cmb_CtaCtb)
End Sub

Private Sub cmb_ConCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ConCtb_Click
   End If
End Sub

Private Sub cmb_CtaBan_Click()
   Call gs_SetFocus(cmb_TipMon)
End Sub

Private Sub cmb_CtaBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaBan_Click
   End If
End Sub

Private Sub cmb_CtaCtb_Change()
   l_str_CtaCtb = cmb_CtaCtb.Text
   
   cmb_CtaCtb.SelLength = Len(l_str_CtaCtb)
End Sub

Private Sub cmb_CtaCtb_Click()
   If cmb_CtaCtb.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmd_Grabar)
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
      
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_CtaCtb_LostFocus()
   Call SendMessage(cmb_CtaCtb.hWnd, CB_SHOWDROPDOWN, 0, 0&)
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
   If cmb_CtaBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Cuenta Bancaria.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaBan)
      Exit Sub
   End If
   
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
   
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTABCO WHERE "
      g_str_Parame = g_str_Parame & "CTABCO_CODBAN = '" & moddat_g_str_CodGrp & "' AND "
      g_str_Parame = g_str_Parame & "CTABCO_NUMCTA = '" & l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "CTABCO_TIPMON = " & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & " AND "
      g_str_Parame = g_str_Parame & "CTABCO_CONCTB = '" & l_arr_ConCtb(cmb_ConCtb.ListIndex + 1).Genera_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "CTABCO_EMPGRP = '" & moddat_g_str_CodEmp & "' "
   
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
      g_str_Parame = "USP_CTB_CTABCO ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & l_arr_ConCtb(cmb_ConCtb.ListIndex + 1).Genera_Codigo & "', "
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
   
   pnl_CodBan.Caption = moddat_g_str_DesGrp
   pnl_EmpGrp.Caption = moddat_g_str_RazSoc
   
   Call fs_Inicio

   Call gs_CentraForm(Me)
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = "SELECT * FROM CTB_CTABCO WHERE "
      g_str_Parame = g_str_Parame & "CTABCO_CODBAN = '" & moddat_g_str_CodGrp & "' AND "
      g_str_Parame = g_str_Parame & "CTABCO_NUMCTA = '" & moddat_g_str_CodIte & "' AND "
      g_str_Parame = g_str_Parame & "CTABCO_TIPMON = " & CStr(moddat_g_int_TipMon) & " AND "
      g_str_Parame = g_str_Parame & "CTABCO_CONCTB = '" & moddat_g_str_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "CTABCO_EMPGRP = '" & moddat_g_str_CodEmp & "'"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         cmb_CtaBan.ListIndex = gf_Busca_Arregl(l_arr_CtaBan, Trim(g_rst_Princi!CTABCO_NUMCTA)) - 1
         
         Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!CTABCO_TIPMON)
         
         cmb_ConCtb.ListIndex = gf_Busca_Arregl(l_arr_ConCtb, Trim(g_rst_Princi!CTABCO_CONCTB)) - 1
         cmb_CtaCtb.ListIndex = gf_Busca_Arregl(l_arr_CtaCtb, Trim(g_rst_Princi!CTABCO_CTACTB)) - 1
         
         cmb_CtaBan.Enabled = False
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
   Call moddat_gs_Carga_CtaBan(moddat_g_str_CodGrp, cmb_CtaBan, l_arr_CtaBan)

   l_int_TopNiv = -1
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, moddat_g_str_CodEmp, "100", "001") Then
      l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
   End If
   
   Call moddat_gs_Carga_CtaCtb(moddat_g_str_CodEmp, cmb_CtaCtb, l_arr_CtaCtb, 0, l_int_TopNiv, -1)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
   Call moddat_gs_Carga_ConceptoCtb(l_arr_ConCtb, cmb_ConCtb, 3)
End Sub


