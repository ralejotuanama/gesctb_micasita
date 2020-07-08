VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_PlaCta_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9780
   ClientLeft      =   4515
   ClientTop       =   1035
   ClientWidth     =   9900
   Icon            =   "GesCtb_frm_139.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9765
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9885
      _Version        =   65536
      _ExtentX        =   17436
      _ExtentY        =   17224
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   9270
         Width           =   9795
         _Version        =   65536
         _ExtentX        =   17277
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
         Begin VB.TextBox txt_BusCta 
            Height          =   315
            Left            =   1380
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "1"
            Top             =   60
            Width           =   3435
         End
         Begin VB.Label Label2 
            Caption         =   "Buscar Cuenta:"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   7275
         Left            =   30
         TabIndex        =   15
         Top             =   1950
         Width           =   9795
         _Version        =   65536
         _ExtentX        =   17277
         _ExtentY        =   12832
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
         Begin MSComctlLib.TreeView tvw_LisCta 
            Height          =   6825
            Left            =   60
            TabIndex        =   7
            Top             =   390
            Width           =   9645
            _ExtentX        =   17013
            _ExtentY        =   12039
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   6
            FullRowSelect   =   -1  'True
            Appearance      =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   9645
            _Version        =   65536
            _ExtentX        =   17013
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuentas Contables"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   13
         Top             =   1470
         Width           =   9795
         _Version        =   65536
         _ExtentX        =   17277
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
         Begin VB.ComboBox cmb_EmpGrp 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   8355
         End
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   10
         Top             =   60
         Width           =   9795
         _Version        =   65536
         _ExtentX        =   17277
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
            TabIndex        =   11
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
            Picture         =   "GesCtb_frm_139.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   12
         Top             =   780
         Width           =   9795
         _Version        =   65536
         _ExtentX        =   17277
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_139.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_139.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_139.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_139.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9180
            Picture         =   "GesCtb_frm_139.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_139.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_PlaCta_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_EmpGrp()      As moddat_tpo_Genera
Dim l_arr_ParEmp()      As moddat_tpo_Genera
Dim l_int_TopNiv        As Integer


Private Sub cmb_EmpGrp_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_EmpGrp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpGrp_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_FlgAct = 1
   
   frm_Mnt_PlaCta_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_EmpGrp.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EmpGrp)
      Exit Sub
   End If

   moddat_g_str_CodGrp = l_arr_EmpGrp(cmb_EmpGrp.ListIndex + 1).Genera_Codigo
   moddat_g_str_DesGrp = cmb_EmpGrp.Text

   l_int_TopNiv = -1
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, moddat_g_str_CodGrp, "100", "001") Then
      l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
   End If
   
   txt_BusCta.MaxLength = l_int_TopNiv
   
   Call fs_Activa(False)
   Call fs_Buscar
End Sub

Private Sub cmd_Editar_Click()
   Dim r_arr_ParEmp()      As moddat_tpo_Genera
   Dim r_int_TopNiv        As Integer

   If tvw_LisCta.Nodes.Count = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   
   frm_Mnt_PlaCta_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_SetFocus(cmb_EmpGrp)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt

   Call fs_Inicio
   Call fs_Limpia
   Call fs_Activa(True)

   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_EmpGrp(cmb_EmpGrp, l_arr_EmpGrp)
End Sub

Private Sub fs_Limpia()
   cmb_EmpGrp.ListIndex = -1
   
   tvw_LisCta.Nodes.Clear
   
   txt_BusCta.Text = ""
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_EmpGrp.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
   tvw_LisCta.Enabled = Not p_Activa
   txt_BusCta.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   Dim r_obj_NodNvo     As Node
   Dim r_int_LarCta     As Integer
   Dim r_lng_NumIte     As Long

   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   tvw_LisCta.Enabled = False
   txt_BusCta.Enabled = False
   
   tvw_LisCta.Nodes.Clear
   
   g_str_Parame = "SELECT * FROM CTB_CTAMAE WHERE "
   g_str_Parame = g_str_Parame & "CTAMAE_CODEMP = '" & moddat_g_str_CodGrp & "' "
   g_str_Parame = g_str_Parame & "ORDER BY CTAMAE_CODCTA ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   r_lng_NumIte = 0
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      If g_rst_Princi!CTAMAE_CODNIV = 1 Then
         If Mid(Trim(g_rst_Princi!CTAMAE_CODCTA), 3, 1) = 0 Then
            Set r_obj_NodNvo = tvw_LisCta.Nodes.Add(, , "C" + Trim(g_rst_Princi!CTAMAE_CODCTA), Trim(g_rst_Princi!CTAMAE_CODCTA) & " - " & Trim(g_rst_Princi!CTAMAE_DESCRI))
            
            r_lng_NumIte = r_lng_NumIte + 1
            tvw_LisCta.Nodes.Item(r_lng_NumIte).ForeColor = modgen_g_con_ColRoj
            'r_obj_NodNvo.Expanded = True
         End If
      Else
         If g_rst_Princi!CTAMAE_CODNIV = 2 Then
            r_int_LarCta = 1
         Else
            r_int_LarCta = g_rst_Princi!CTAMAE_CODNIV - 2
         End If
      
         Set r_obj_NodNvo = tvw_LisCta.Nodes.Add("C" + Mid(Trim(g_rst_Princi!CTAMAE_CODCTA), 1, r_int_LarCta) & String(l_int_TopNiv - r_int_LarCta, "0"), tvwChild, "C" + Trim(g_rst_Princi!CTAMAE_CODCTA), Trim(g_rst_Princi!CTAMAE_CODCTA) & " - " & Trim(g_rst_Princi!CTAMAE_DESCRI))
         r_lng_NumIte = r_lng_NumIte + 1
         
         Select Case g_rst_Princi!CTAMAE_CODNIV
            Case 2, 14:    tvw_LisCta.Nodes.Item(r_lng_NumIte).ForeColor = modgen_g_con_ColAzu
            Case 4, 16:    tvw_LisCta.Nodes.Item(r_lng_NumIte).ForeColor = modgen_g_con_ColMag
            Case 6, 18:    tvw_LisCta.Nodes.Item(r_lng_NumIte).ForeColor = modgen_g_con_ColVer
            Case 8, 20:    tvw_LisCta.Nodes.Item(r_lng_NumIte).ForeColor = modgen_g_con_ColNar
            Case 10, 22:   tvw_LisCta.Nodes.Item(r_lng_NumIte).ForeColor = modgen_g_con_ColCya
            Case 12, 24:   tvw_LisCta.Nodes.Item(r_lng_NumIte).ForeColor = modgen_g_con_ColNeg
         End Select
         
         'r_obj_NodNvo.EnsureVisible
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'tvw_LisCta.SelectedItem
   r_obj_NodNvo.Root.Selected = True
   
   
   cmd_Editar.Enabled = True
   cmd_Borrar.Enabled = True
   tvw_LisCta.Enabled = True
   txt_BusCta.Enabled = True
   
   Call gs_SetFocus(tvw_LisCta)
End Sub


Private Sub tvw_LisCta_DblClick()
   'Call cmd_Editar_Click
End Sub

Private Sub tvw_LisCta_NodeClick(ByVal Node As MSComctlLib.Node)
   'Node.LastSibling.Expanded = False
   
   moddat_g_str_Codigo = Mid(Node.Text, 1, l_int_TopNiv)
   moddat_g_str_Descri = Mid(Node.Text, l_int_TopNiv + 3)
End Sub

Private Sub txt_BusCta_KeyPress(KeyAscii As Integer)
   Dim r_obj_NodNvo     As Node
   Dim r_lng_Contad     As Long
   Dim r_int_FlgEnc     As Integer
   
   If KeyAscii = 13 Then
      If Len(Trim(txt_BusCta.Text)) = 0 Then
         Exit Sub
      End If
      
      r_int_FlgEnc = 0
   
      For r_lng_Contad = 1 To tvw_LisCta.Nodes.Count
         If Mid(tvw_LisCta.Nodes(r_lng_Contad).Key, 1, Len(txt_BusCta.Text) + 1) = "C" & txt_BusCta.Text Then
            tvw_LisCta.Nodes.Item(r_lng_Contad).Selected = True
            
            r_int_FlgEnc = 1
            
            Call gs_SetFocus(tvw_LisCta)
            Exit For
         End If
      Next r_lng_Contad
      
      If r_int_FlgEnc = 0 Then
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub
