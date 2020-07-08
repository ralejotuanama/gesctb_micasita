VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Mnt_Bancos_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4275
   ClientLeft      =   2775
   ClientTop       =   4800
   ClientWidth     =   12390
   Icon            =   "GesCtb_frm_158.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4275
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12405
      _Version        =   65536
      _ExtentX        =   21881
      _ExtentY        =   7541
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
         TabIndex        =   9
         Top             =   60
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
            TabIndex        =   10
            Top             =   60
            Width           =   5085
            _Version        =   65536
            _ExtentX        =   8969
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de Bancos"
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
            Picture         =   "GesCtb_frm_158.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   11
         Top             =   780
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
            Left            =   30
            Picture         =   "GesCtb_frm_158.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11670
            Picture         =   "GesCtb_frm_158.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   435
         Left            =   30
         TabIndex        =   12
         Top             =   1470
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
         Begin Threed.SSPanel pnl_Bancos 
            Height          =   315
            Left            =   1470
            TabIndex        =   13
            Top             =   60
            Width           =   10755
            _Version        =   65536
            _ExtentX        =   18971
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
            Height          =   285
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1785
         Left            =   30
         TabIndex        =   15
         Top             =   1950
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
         _ExtentY        =   3149
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
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1470
            MaxLength       =   250
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1080
            Width           =   10755
         End
         Begin VB.TextBox txt_NumCta 
            Height          =   315
            Left            =   1470
            MaxLength       =   25
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   90
            Width           =   3225
         End
         Begin VB.ComboBox cmb_TipCta 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   3225
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   750
            Width           =   3225
         End
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1410
            Width           =   3225
         End
         Begin VB.Label Label7 
            Caption         =   "Descripción:"
            Height          =   285
            Left            =   60
            TabIndex        =   20
            Top             =   1080
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Moneda:"
            Height          =   285
            Left            =   60
            TabIndex        =   19
            Top             =   750
            Width           =   1305
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   90
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Cuenta:"
            Height          =   285
            Left            =   60
            TabIndex        =   17
            Top             =   420
            Width           =   1245
         End
         Begin VB.Label Label6 
            Caption         =   "Situación:"
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   1410
            Width           =   1275
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   21
         Top             =   3780
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
         Begin VB.ComboBox cmb_CtaCtb 
            Height          =   315
            Left            =   1470
            TabIndex        =   5
            Text            =   "cmb_CtaCtb"
            Top             =   60
            Width           =   10755
         End
         Begin VB.Label Label8 
            Caption         =   "Cuenta Contable:"
            Height          =   285
            Left            =   60
            TabIndex        =   22
            Top             =   90
            Width           =   1515
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_Bancos_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_TipCta()      As moddat_tpo_Genera
Dim l_arr_ParEmp()      As moddat_tpo_Genera
Dim l_arr_CtaCtb()      As moddat_tpo_Genera
Dim l_str_CtaCtb        As String
Dim l_int_FlgCmb        As Integer
Dim l_int_TopNiv        As Integer


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

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(cmb_CtaCtb)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub cmb_TipCta_Click()
   Call gs_SetFocus(cmb_TipMon)
End Sub

Private Sub cmb_TipCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipCta_Click
   End If
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_NumCta.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumCta)
      Exit Sub
   End If
   
   If cmb_TipCta.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCta)
      Exit Sub
   End If
   
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If

   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar si la Cuenta Bancaria es Vigente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If

   If cmb_CtaCtb.ListIndex = -1 Then
      MsgBox "Debe seleccionar una Cuenta Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaCtb)
      Exit Sub
   End If

   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM MNT_CTABAN WHERE "
      g_str_Parame = g_str_Parame & "CTABAN_CODBAN = '" & moddat_g_str_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "CTABAN_NUMCTA = '" & txt_NumCta.Text & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
        
         MsgBox "La Cuenta ya ha sido registrada. Por favor verifique el número e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_MNT_CTABAN ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumCta.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_TipCta(cmb_TipCta.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & l_arr_CtaCtb(cmb_CtaCtb.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
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

Private Sub txt_Descri_GotFocus()
   Call gs_SelecTodo(txt_Descri)
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Situac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_ .,;:()=&/%$@#")
   End If
End Sub

Private Sub txt_NumCta_GotFocus()
   Call gs_SelecTodo(txt_NumCta)
End Sub

Private Sub txt_NumCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipCta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   
   pnl_Bancos.Caption = moddat_g_str_DesGrp
   
   Call fs_Inicia
   Call fs_Limpia
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = "SELECT * FROM MNT_CTABAN WHERE CTABAN_CODBAN = '" & moddat_g_str_CodGrp & "' AND CTABAN_NUMCTA = '" & moddat_g_str_Codigo & "' "
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         txt_NumCta.Text = moddat_g_str_Codigo
         txt_NumCta.Enabled = False
         
         cmb_TipCta.ListIndex = gf_Busca_Arregl(l_arr_TipCta, Trim(g_rst_Princi!CtaBan_TipCta)) - 1
         
         Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!ctaban_TipMon)
         
         txt_Descri.Text = Trim(g_rst_Princi!ctaban_Descri & "")
         
         Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Princi!ctaban_Situac)
         
         cmb_CtaCtb.ListIndex = gf_Busca_Arregl(l_arr_CtaCtb, Trim(g_rst_Princi!CtaBan_CtaCtb & "")) - 1
         
         Call gs_SetFocus(cmb_TipCta)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_TipCta, l_arr_TipCta, 1, "510")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")

   l_int_TopNiv = -1
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, "000001", "100", "001") Then
      l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
   End If
   
   Call moddat_gs_Carga_CtaCtb("000001", cmb_CtaCtb, l_arr_CtaCtb, 0, l_int_TopNiv, -1)
End Sub

Private Sub fs_Limpia()
   txt_NumCta.Text = ""
   cmb_TipCta.ListIndex = -1
   cmb_TipMon.ListIndex = -1
   txt_Descri.Text = ""
   cmb_Situac.ListIndex = -1
   cmb_CtaCtb.ListIndex = -1
End Sub


