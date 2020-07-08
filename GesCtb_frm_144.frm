VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_PlaCta_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   4740
   ClientLeft      =   11985
   ClientTop       =   3045
   ClientWidth     =   9840
   Icon            =   "GesCtb_frm_144.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   4725
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9825
      _Version        =   65536
      _ExtentX        =   17330
      _ExtentY        =   8334
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
         Height          =   1095
         Left            =   30
         TabIndex        =   22
         Top             =   3090
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
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
         Begin VB.ComboBox cmb_RegCom 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   3405
         End
         Begin VB.ComboBox cmb_TraSal 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   390
            Width           =   3405
         End
         Begin VB.ComboBox cmb_NivCta 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   60
            Width           =   3405
         End
         Begin VB.Label Label4 
            Caption         =   "Registro Comprob.:"
            Height          =   285
            Left            =   60
            TabIndex        =   25
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Saldo al:"
            Height          =   285
            Left            =   60
            TabIndex        =   24
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Nivel Cuenta:"
            Height          =   285
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   1155
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   20
         Top             =   4230
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
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
         Begin VB.ComboBox cmb_Situac 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   60
            Width           =   3405
         End
         Begin VB.Label Label6 
            Caption         =   "Situación:"
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   1155
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   1095
         Left            =   30
         TabIndex        =   10
         Top             =   1950
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
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
         Begin VB.TextBox txt_DesCor 
            Height          =   315
            Left            =   1560
            MaxLength       =   60
            TabIndex        =   2
            Text            =   "1"
            Top             =   720
            Width           =   8115
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1560
            MaxLength       =   250
            TabIndex        =   1
            Text            =   "1"
            Top             =   390
            Width           =   8115
         End
         Begin VB.TextBox txt_CodCta 
            Height          =   315
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   0
            Text            =   "1"
            Top             =   60
            Width           =   3435
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Abreviatura:"
            Height          =   285
            Index           =   2
            Left            =   60
            TabIndex        =   19
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Descripción:"
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   18
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Cuenta:"
            Height          =   285
            Index           =   1
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1185
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   11
         Top             =   1470
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
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
         Begin Threed.SSPanel pnl_NomEmp 
            Height          =   315
            Left            =   1560
            TabIndex        =   17
            Top             =   60
            Width           =   8115
            _Version        =   65536
            _ExtentX        =   14314
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   13
         Top             =   60
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
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
            TabIndex        =   14
            Top             =   60
            Width           =   5085
            _Version        =   65536
            _ExtentX        =   8969
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de Plan de Cuentas"
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
            Picture         =   "GesCtb_frm_144.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   15
         Top             =   780
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
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
            Picture         =   "GesCtb_frm_144.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9120
            Picture         =   "GesCtb_frm_144.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_PlaCta_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ParEmp()      As moddat_tpo_Genera
Dim l_int_TopNiv        As Integer

Private Sub cmd_Grabar_Click()
   Dim r_str_DigMon     As String
   Dim r_str_Descri     As String
   Dim r_int_LarCta     As Integer

   If Len(Trim(txt_CodCta.Text)) <> l_int_TopNiv Then
      MsgBox "La cuenta debe tener la longitud del Tope de Nivel de Cuenta para esta empresa (" & CStr(l_int_TopNiv) & ").", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_CodCta)
      Exit Sub
   End If
   
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "La Descripción de la Cuenta está vacía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If

   If Len(Trim(txt_DesCor.Text)) = 0 Then
      MsgBox "La Abreviatura de la Cuenta está vacía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DesCor)
      Exit Sub
   End If
   
   If cmb_NivCta.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Nivel de Cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NivCta)
      Exit Sub
   End If
   
   If cmb_TraSal.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tratamiento de Saldo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TraSal)
      Exit Sub
   End If
   
   If cmb_RegCom.ListIndex = -1 Then
      MsgBox "Debe seleccionar si se permite el Registro de Comprobantes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_RegCom)
      Exit Sub
   End If
   
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   
   'Validar que cuenta no tenga Registro de Comprobantes en Niveles Superiores
   If cmb_NivCta.ItemData(cmb_NivCta.ListIndex) > 1 Then
      r_int_LarCta = 0
      
      If cmb_NivCta.ItemData(cmb_NivCta.ListIndex) = 2 Then
         r_int_LarCta = 1
      Else
         r_int_LarCta = cmb_NivCta.ItemData(cmb_NivCta.ListIndex) - 2
      End If
      
      g_str_Parame = "SELECT * FROM CTB_CTAMAE WHERE "
      g_str_Parame = g_str_Parame & "CTAMAE_CODEMP = '" & moddat_g_str_CodGrp & "' AND "
      g_str_Parame = g_str_Parame & "CTAMAE_CODCTA = '" & Mid(txt_CodCta.Text, 1, r_int_LarCta) & String(l_int_TopNiv - r_int_LarCta, "0") & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing

         MsgBox "La Cuenta Contable no puede ser registrada porque no existe Cuenta de Nivel Superior.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_NivCta)
         Exit Sub
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = "SELECT * FROM CTB_CTAMAE WHERE CTAMAE_CODEMP = '" & moddat_g_str_CodGrp & "' AND CTAMAE_CODCTA = '" & txt_CodCta.Text & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing

         MsgBox "La Cuenta Contable ya ha sido registrada.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_CodCta)
         Exit Sub
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Call fs_Grabar(txt_CodCta.Text, txt_Descri.Text, txt_DesCor.Text)
   
   If moddat_g_int_FlgGrb = 1 And Mid(txt_CodCta.Text, 3, 1) = "0" Then
      g_str_Parame = "SELECT * FROM MNT_PAREMP WHERE PAREMP_CODEMP = '" & moddat_g_str_CodGrp & "' AND PAREMP_CODGRP = '101' AND PAREMP_CODITE <> '000000' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
         
         Do While Not g_rst_Genera.EOF
            r_str_DigMon = CStr(CInt(g_rst_Genera!PAREMP_CODITE))
            r_str_Descri = Trim(g_rst_Genera!PAREMP_DESCRI)
                        
            If MsgBox("Ha creado una cuenta en Dígito Integrador. ¿Desea crear la cuenta en Dígito " & r_str_DigMon & "?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
               g_str_Parame = "SELECT * FROM CTB_CTAMAE WHERE CTAMAE_CODEMP = '" & moddat_g_str_CodGrp & "' AND CTAMAE_CODCTA = '" & Mid(txt_CodCta.Text, 1, 2) & r_str_DigMon & Mid(txt_CodCta, 4) & "' "
               
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
                   Exit Sub
               End If
            
               If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
                  g_rst_Princi.Close
                  Set g_rst_Princi = Nothing
         
                  MsgBox "La Cuenta Contable ya se encuentra registrada.", vbExclamation, modgen_g_str_NomPlt
               Else
                  g_rst_Princi.Close
                  Set g_rst_Princi = Nothing
                  
                  Call fs_Grabar(Mid(txt_CodCta.Text, 1, 2) & r_str_DigMon & Mid(txt_CodCta, 4), Mid(txt_Descri.Text & " (" & r_str_Descri & ")", 1, 250), Mid(txt_DesCor.Text & " (" & r_str_Descri & ")", 1, 60))
               End If
            End If
         
            g_rst_Genera.MoveNext
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_int_Contad     As Integer
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   
   pnl_NomEmp.Caption = moddat_g_str_CodGrp & " - " & moddat_g_str_DesGrp
   
   Call fs_Inicio
   Call fs_Limpia
   
   l_int_TopNiv = -1
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, moddat_g_str_CodGrp, "100", "001") Then
      l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
   End If
   
   If l_int_TopNiv = -1 Then
      Screen.MousePointer = 0
      
      MsgBox "No se ha encontrado el registro de Tope de Nivel de Cuenta para esta empresa.", vbExclamation, modgen_g_str_NomPlt
      
      cmd_Grabar.Enabled = False
      
      txt_CodCta.Enabled = False
      txt_Descri.Enabled = False
      txt_DesCor.Enabled = False
      
      cmb_NivCta.Enabled = False
      cmb_TraSal.Enabled = False
      cmb_RegCom.Enabled = False
      
      cmb_Situac.Enabled = False
      
      Exit Sub
   End If
   
   
   For r_int_Contad = cmb_NivCta.ListCount - 1 To 1 Step -1
      If cmb_NivCta.ItemData(r_int_Contad) > l_int_TopNiv Then
         cmb_NivCta.RemoveItem r_int_Contad
      End If
   Next r_int_Contad
   
   txt_CodCta.MaxLength = l_int_TopNiv
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = "SELECT * FROM CTB_CTAMAE WHERE CTAMAE_CODEMP = '" & moddat_g_str_CodGrp & "' AND CTAMAE_CODCTA = '" & moddat_g_str_Codigo & "' "
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         txt_CodCta.Text = Trim(g_rst_Princi!CTAMAE_CODCTA)
         
         txt_CodCta.Enabled = False
         
         txt_Descri.Text = Trim(g_rst_Princi!CTAMAE_DESCRI)
         txt_DesCor.Text = Trim(g_rst_Princi!CTAMAE_DESCOR)
         
         Call gs_BuscarCombo_Item(cmb_NivCta, g_rst_Princi!CTAMAE_CODNIV)
         Call gs_BuscarCombo_Item(cmb_TraSal, g_rst_Princi!CTAMAE_TRASAL)
         Call gs_BuscarCombo_Item(cmb_RegCom, g_rst_Princi!CTAMAE_REGCOM)
         Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Princi!CTAMAE_SITUAC)
         
         Call gs_SetFocus(txt_Descri)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_LisIte_Combo(cmb_NivCta, 1, "256")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TraSal, 1, "255")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
   Call moddat_gs_Carga_LisIte_Combo(cmb_RegCom, 1, "214")
End Sub
   
Private Sub fs_Limpia()
   txt_CodCta.Text = ""
   txt_Descri.Text = ""
   txt_DesCor.Text = ""
   
   cmb_NivCta.ListIndex = -1
   cmb_TraSal.ListIndex = -1
   cmb_RegCom.ListIndex = -1
   
   cmb_Situac.ListIndex = -1
End Sub

Private Sub txt_CodCta_GotFocus()
   Call gs_SelecTodo(txt_CodCta)
End Sub

Private Sub txt_CodCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descri)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_CodCta_LostFocus()
   If Len(Trim(txt_CodCta.Text)) > 0 Then
      txt_CodCta.Text = txt_CodCta.Text & String(l_int_TopNiv - Len(txt_CodCta.Text), "0")
   End If
End Sub

Private Sub txt_Descri_GotFocus()
   Call gs_SelecTodo(txt_Descri)
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DesCor)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:()=/#")
   End If
End Sub

Private Sub txt_DesCor_GotFocus()
   Call gs_SelecTodo(txt_DesCor)
End Sub

Private Sub txt_DesCor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_NivCta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:()=/#")
   End If
End Sub

Private Sub cmb_NivCta_Click()
   Call gs_SetFocus(cmb_TraSal)
End Sub

Private Sub cmb_NivCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NivCta_Click
   End If
End Sub

Private Sub cmb_TraSal_Click()
   Call gs_SetFocus(cmb_RegCom)
End Sub

Private Sub cmb_TraSal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TraSal_Click
   End If
End Sub

Private Sub cmb_RegCom_Click()
   Call gs_SetFocus(cmb_Situac)
End Sub

Private Sub cmb_RegCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_RegCom_Click
   End If
End Sub

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub fs_Grabar(ByVal p_CodCta As String, ByVal p_Descri As String, ByVal p_DesCor As String)
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CTB_CTAMAE ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
      g_str_Parame = g_str_Parame & "'" & p_CodCta & "', "
      g_str_Parame = g_str_Parame & "'" & p_Descri & "', "
      g_str_Parame = g_str_Parame & "'" & p_DesCor & "', "
      
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_NivCta.ItemData(cmb_NivCta.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TraSal.ItemData(cmb_TraSal.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_RegCom.ItemData(cmb_RegCom.ListIndex)) & ", "
         
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
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
End Sub

Private Sub txt_Descri_LostFocus()
   txt_DesCor.Text = txt_Descri.Text
End Sub
