VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_EmpSup_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3300
   ClientLeft      =   12540
   ClientTop       =   5055
   ClientWidth     =   7470
   Icon            =   "GesCtb_frm_164.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3285
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   5794
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
         TabIndex        =   8
         Top             =   60
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
            TabIndex        =   9
            Top             =   60
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de Empresas Supervisadas"
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
            Picture         =   "GesCtb_frm_164.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1755
         Left            =   30
         TabIndex        =   10
         Top             =   1470
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   3096
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
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1380
            Width           =   5445
         End
         Begin VB.ComboBox cmb_TipEnt 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1050
            Width           =   5445
         End
         Begin VB.TextBox txt_NomCor 
            Height          =   315
            Left            =   1860
            MaxLength       =   250
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   720
            Width           =   5445
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1860
            MaxLength       =   5
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   1305
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1860
            MaxLength       =   250
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   5445
         End
         Begin VB.Label Label4 
            Caption         =   "Situación:"
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   1380
            Width           =   1665
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Entidad:"
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   1050
            Width           =   1665
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre Corto:"
            Height          =   285
            Left            =   60
            TabIndex        =   14
            Top             =   720
            Width           =   1665
         End
         Begin VB.Label Label8 
            Caption         =   "Código Empresa:"
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre Empresa:"
            Height          =   285
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   1665
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   13
         Top             =   780
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
            Left            =   6750
            Picture         =   "GesCtb_frm_164.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_164.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_EmpSup_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_Situac_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_Situac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Situac_Click
   End If
End Sub

Private Sub cmb_TipEnt_Click()
   Call gs_SetFocus(cmb_Situac)
End Sub

Private Sub cmb_TipEnt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipEnt_Click
   End If
End Sub

Private Sub txt_Codigo_GotFocus()
   Call gs_SelecTodo(txt_Codigo)
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomCor)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,;:()")
   End If
End Sub

Private Sub txt_NomCor_GotFocus()
   Call gs_SelecTodo(txt_NomCor)
End Sub

Private Sub txt_NomCor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipEnt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ .,;:()")
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_Codigo.Text)) = 0 Then
      MsgBox "Debe ingresar el Código de Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Codigo)
      Exit Sub
   End If
   
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If
   
   If Len(Trim(txt_NomCor.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre Corto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomCor)
      Exit Sub
   End If
   
   If cmb_TipEnt.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Entidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipEnt)
      Exit Sub
   End If
   
   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = "SELECT * FROM CTB_EMPSUP WHERE EMPSUP_CODIGO = " & txt_Codigo.Text & " "
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing

         MsgBox "El Código de Empresa ya ha sido registrado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
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
      g_str_Parame = "USP_CTB_EMPSUP ("
      g_str_Parame = g_str_Parame & txt_Codigo.Text & ", "
      g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NomCor.Text & "', "
      
      g_str_Parame = g_str_Parame & CStr(cmb_TipEnt.ItemData(cmb_TipEnt.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_Situac.ItemData(cmb_Situac.ListIndex)) & ", "
         
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
   
   Call fs_Inicia
   Call fs_Limpia
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = "SELECT * FROM CTB_EMPSUP WHERE EMPSUP_CODIGO = " & moddat_g_str_CodGrp & " "
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         txt_Codigo.Text = moddat_g_str_CodGrp
         txt_Codigo.Enabled = False
         
         txt_Nombre.Text = Trim(g_rst_Princi!EMPSUP_NOMBRE)
         txt_NomCor.Text = Trim(g_rst_Princi!EMPSUP_NOMCOR)
         
         Call gs_BuscarCombo_Item(cmb_TipEnt, g_rst_Princi!EMPSUP_TIPENT)
         Call gs_BuscarCombo_Item(cmb_Situac, g_rst_Princi!EMPSUP_SITUAC)
         
         Call gs_SetFocus(txt_Nombre)
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipEnt, 1, "263")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Situac, 1, "013")
End Sub

Private Sub fs_Limpia()
   txt_Codigo.Text = ""
   txt_Nombre.Text = ""
   txt_NomCor.Text = ""
   cmb_TipEnt.ListIndex = -1
   cmb_Situac.ListIndex = -1
End Sub

