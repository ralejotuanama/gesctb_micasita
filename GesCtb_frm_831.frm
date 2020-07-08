VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_Person_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   3075
   ClientLeft      =   2700
   ClientTop       =   5235
   ClientWidth     =   5400
   Icon            =   "GesCtb_frm_831.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5445
      _Version        =   65536
      _ExtentX        =   9604
      _ExtentY        =   5477
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
         TabIndex        =   7
         Top             =   30
         Width           =   5325
         _Version        =   65536
         _ExtentX        =   9393
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   400
            Left            =   600
            TabIndex        =   8
            Top             =   60
            Width           =   3000
            _Version        =   65536
            _ExtentX        =   5292
            _ExtentY        =   706
            _StockProps     =   15
            Caption         =   "Mantenimiento de Personal"
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
            Picture         =   "GesCtb_frm_831.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   675
         Left            =   30
         TabIndex        =   9
         Top             =   750
         Width           =   5325
         _Version        =   65536
         _ExtentX        =   9393
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_831.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4650
            Picture         =   "GesCtb_frm_831.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1095
         Left            =   30
         TabIndex        =   10
         Top             =   1470
         Width           =   5325
         _Version        =   65536
         _ExtentX        =   9393
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
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   2
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   1
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   3
            Top             =   720
            Width           =   3315
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   60
            TabIndex        =   13
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   60
            TabIndex        =   11
            Top             =   720
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   14
         Top             =   2610
         Width           =   5325
         _Version        =   65536
         _ExtentX        =   9393
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
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label Label14 
            Caption         =   "Categoría"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1665
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_Person_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim r_int_PerMes        As String
   Dim r_int_PerAno        As String

Private Sub cmd_Grabar_Click()

   If Trim(txt_ApePat.Text) = "" Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   
   If Trim(txt_ApeMat.Text) = "" Then
      MsgBox "Debe ingresar el Apellido Materno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApeMat)
      Exit Sub
   End If
   
   If Trim(txt_Nombre.Text) = "" Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If

   If cmb_Situac.ListIndex = -1 Then
      MsgBox "Debe seleccionar una Categoría.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Situac)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro que desea registrar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
  
   
   If moddat_g_int_FlgGrb = 1 Then
   
      g_str_Parame = "INSERT INTO CTB_USUSBS VALUES ("
      g_str_Parame = g_str_Parame & r_int_PerMes & ", "
      g_str_Parame = g_str_Parame & r_int_PerAno & ", "
      g_str_Parame = g_str_Parame & "'" & Left(txt_Nombre.Text, 1) & txt_ApePat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_ApePat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_ApeMat.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Nombre.Text & "', "
      g_str_Parame = g_str_Parame & cmb_Situac.ListIndex + 1 & ") "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      MsgBox "Ingreso exitoso.", vbInformation, modgen_g_str_NomPlt

   ElseIf moddat_g_int_FlgGrb = 2 Then
   
      g_str_Parame = "UPDATE CTB_USUSBS SET "
      g_str_Parame = g_str_Parame & "USUSBS_APEPAT = '" & txt_ApePat.Text & "', "
      g_str_Parame = g_str_Parame & "USUSBS_APEMAT = '" & txt_ApeMat.Text & "', "
      g_str_Parame = g_str_Parame & "USUSBS_NOMBRE = '" & txt_Nombre.Text & "', "
      g_str_Parame = g_str_Parame & "USUSBS_TIPPER = " & cmb_Situac.ListIndex + 1
      g_str_Parame = g_str_Parame & "WHERE USUSBS_CODUSU = '" & modsec_g_str_CodUsu & "' AND "
      g_str_Parame = g_str_Parame & "USUSBS_PERMES = " & r_int_PerMes & " AND USUSBS_PERANO = " & r_int_PerAno & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   
      MsgBox "Modificación exitosa.", vbInformation, modgen_g_str_NomPlt
      
   End If
   
   Call cmd_Salida_Click
   
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt

   Call gs_CentraForm(Me)
   Call fs_Inicio
   Call ff_BusPer
   
   Screen.MousePointer = 0
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Situac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub fs_Inicio()

   cmb_Situac.AddItem "GERENTE"
   cmb_Situac.AddItem "FUNCIONARIO"
   cmb_Situac.AddItem "EMPLEADO"
   cmb_Situac.AddItem "OTROS"
   
   If moddat_g_int_FlgGrb = 2 Then
      Call fs_Buscar
     
   ElseIf moddat_g_int_FlgGrb = 1 Then
      
      txt_ApePat.Text = ""
      txt_ApeMat.Text = ""
      txt_Nombre.Text = ""

      cmb_Situac.ListIndex = -1
   
   End If
   
End Sub

Private Sub fs_Buscar()
   
   g_str_Parame = "SELECT * FROM CTB_USUSBS "
   g_str_Parame = g_str_Parame & "WHERE USUSBS_CODUSU = '" & modsec_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "ORDER BY "
   g_str_Parame = g_str_Parame & "USUSBS_APEPAT, USUSBS_APEMAT ASC, USUSBS_NOMBRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
      
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
      txt_ApePat.Text = Trim(g_rst_Princi!USUSBS_APEPAT)
      txt_ApeMat.Text = Trim(g_rst_Princi!USUSBS_APEMAT)
      txt_Nombre.Text = Trim(g_rst_Princi!USUSBS_NOMBRE)
      cmb_Situac.ListIndex = Trim(g_rst_Princi!USUSBS_TIPPER) - 1
      
      g_rst_Princi.MoveNext
   Loop
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
End Sub

Public Sub ff_BusPer()
      
   g_str_Parame = "SELECT * FROM CTB_USUSBS ORDER BY USUSBS_PERANO DESC, USUSBS_PERMES DESC"
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
       r_int_PerMes = g_rst_Listas!USUSBS_PERMES
       r_int_PerAno = g_rst_Listas!USUSBS_PERANO
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub
