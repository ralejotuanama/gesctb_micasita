VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_DetGar_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   34305
   ClientTop       =   8190
   ClientWidth     =   7470
   Icon            =   "GesCtb_frm_125.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3285
      Left            =   0
      TabIndex        =   5
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   6
         Top             =   2790
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
         Begin VB.ComboBox cmb_ClaGar 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   5445
         End
         Begin VB.Label Label2 
            Caption         =   "Clase Garant�a:"
            Height          =   285
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   1485
         End
      End
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
            Width           =   5085
            _Version        =   65536
            _ExtentX        =   8969
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clasificaci�n por Tipo de Cr�dito"
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
            Picture         =   "GesCtb_frm_125.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   795
         Left            =   30
         TabIndex        =   10
         Top             =   1950
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1860
            MaxLength       =   250
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   5445
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1860
            MaxLength       =   3
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "Descripci�n:"
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "C�digo Garant�a:"
            Height          =   285
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1485
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_125.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6750
            Picture         =   "GesCtb_frm_125.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   435
         Left            =   30
         TabIndex        =   14
         Top             =   1470
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
         Begin Threed.SSPanel pnl_TipGar 
            Height          =   315
            Left            =   1860
            TabIndex        =   15
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
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
         Begin VB.Label Label4 
            Caption         =   "Tipo de Garant�a:"
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_DetGar_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ClaGar()      As moddat_tpo_Genera

Private Sub cmb_ClaGar_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_ClaGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ClaGar_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_Codigo.Text)) = 0 Then
      MsgBox "Debe ingresar el C�digo de Garant�a.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Codigo)
      Exit Sub
   End If
   
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripci�n.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If

   If cmb_ClaGar.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Clase de Garant�a.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ClaGar)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = "SELECT * FROM CTB_DETGAR WHERE DETGAR_CODIGO = '" & txt_Codigo.Text & "' "
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing

         MsgBox "El C�digo de Garant�a ya ha sido registrado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Codigo)
         Exit Sub
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   If MsgBox("�Est� seguro de grabar los datos?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CTB_DETGAR ("
      g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_ClaGar(cmb_ClaGar.ListIndex + 1).Genera_Codigo & "', "
         
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
         If MsgBox("No se pudo completar el procedimiento. �Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
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
   
   pnl_TipGar.Caption = moddat_g_str_Codigo & " - " & moddat_g_str_Descri
   
   Call fs_Inicia
   Call fs_Limpia
   
   If moddat_g_int_FlgGrb = 2 Then
      g_str_Parame = "SELECT * FROM CTB_DETGAR WHERE DETGAR_CODIGO = '" & moddat_g_str_CodIte & "' "
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         txt_Codigo.Text = moddat_g_str_CodIte
         txt_Codigo.Enabled = False
         
         txt_Descri.Text = Trim(g_rst_Princi!DETGAR_DESCRI)
         
         cmb_ClaGar.ListIndex = gf_Busca_Arregl(l_arr_ClaGar, Trim(g_rst_Princi!DetGar_ClaGar)) - 1
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Screen.MousePointer = 0
End Sub
   
Private Sub fs_Inicia()
   Call moddat_gs_Carga_ClaGar(cmb_ClaGar, l_arr_ClaGar)
End Sub

Private Sub fs_Limpia()
   txt_Codigo.Text = ""
   txt_Descri.Text = ""
   
   cmb_ClaGar.ListIndex = -1
End Sub

Private Sub txt_Codigo_GotFocus()
   Call gs_SelecTodo(txt_Codigo)
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descri)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Descri_GotFocus()
   Call gs_SelecTodo(txt_Descri)
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ClaGar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO + modgen_g_con_LETRAS + ", .-_;:)(=?�/&%$")
   End If
End Sub


