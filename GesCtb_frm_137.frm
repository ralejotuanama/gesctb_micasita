VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_Provis_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   5400
   ClientTop       =   2355
   ClientWidth     =   9900
   Icon            =   "GesCtb_frm_137.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7125
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9885
      _Version        =   65536
      _ExtentX        =   17436
      _ExtentY        =   12568
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
         TabIndex        =   11
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
            TabIndex        =   12
            Top             =   60
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de Provisiones"
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
            Picture         =   "GesCtb_frm_137.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   4785
         Left            =   30
         TabIndex        =   13
         Top             =   2280
         Width           =   9795
         _Version        =   65536
         _ExtentX        =   17277
         _ExtentY        =   8440
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   4395
            Left            =   30
            TabIndex        =   8
            Top             =   360
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   7752
            _Version        =   393216
            Rows            =   25
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   4140
            TabIndex        =   14
            Top             =   60
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Clase Garantía"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   4095
            _Version        =   65536
            _ExtentX        =   7223
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Clasificación Crediticia"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   8220
            TabIndex        =   16
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "%"
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   9
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_137.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9180
            Picture         =   "GesCtb_frm_137.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_137.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_137.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_137.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_137.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   17
         Top             =   1470
         Width           =   9795
         _Version        =   65536
         _ExtentX        =   17277
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
         Begin VB.ComboBox cmb_ClaCre 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   8055
         End
         Begin VB.ComboBox cmb_TipPrv 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   8055
         End
         Begin VB.Label Label1 
            Caption         =   "Clase de Crédito:"
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo Provisión:"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_Provis_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ClaCre()      As moddat_tpo_Genera

Private Sub cmb_ClaCre_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_ClaCre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ClaCre_Click
   End If
End Sub

Private Sub cmb_TipPrv_Click()
   Call gs_SetFocus(cmb_ClaCre)
End Sub

Private Sub cmb_TipPrv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPrv_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_FlgAct = 1
   
   frm_Mnt_Provis_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 3
   moddat_g_str_CodIte = grd_Listad.Text
   
   grd_Listad.Col = 4
   moddat_g_str_CodMod = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = "DELETE FROM CTB_TIPPRV WHERE "
   g_str_Parame = g_str_Parame & "TIPPRV_TIPPRV = '" & moddat_g_str_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "TIPPRV_CLACRE = '" & moddat_g_str_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "TIPPRV_CLFCRE = '" & moddat_g_str_CodIte & "' AND "
   g_str_Parame = g_str_Parame & "TIPPRV_CLAGAR = '" & moddat_g_str_CodMod & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0

End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipPrv.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Provisión.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipPrv)
      Exit Sub
   End If
   
   If cmb_ClaCre.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Clase de Crédito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ClaCre)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Call fs_Activa(False)
   Call fs_Buscar
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 3
   moddat_g_str_CodIte = grd_Listad.Text
   
   grd_Listad.Col = 4
   moddat_g_str_CodMod = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   
   frm_Mnt_Provis_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Limpia_Click()
   cmb_TipPrv.ListIndex = -1
   cmb_ClaCre.ListIndex = -1
   
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Activa(True)
   
   Call gs_SetFocus(cmb_TipPrv)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)

   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 4085
   grd_Listad.ColWidth(1) = 4085
   grd_Listad.ColWidth(2) = 1145
   grd_Listad.ColWidth(3) = 0
   grd_Listad.ColWidth(4) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   
   Call moddat_gs_Carga_TipCre(cmb_ClaCre, l_arr_ClaCre)
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPrv, 1, "352")
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_TipPrv.Enabled = p_Activa
   cmb_ClaCre.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   moddat_g_str_Codigo = CStr(cmb_TipPrv.ItemData(cmb_TipPrv.ListIndex))
   moddat_g_str_Descri = cmb_TipPrv.Text
   
   moddat_g_str_CodGrp = l_arr_ClaCre(cmb_ClaCre.ListIndex + 1).Genera_Codigo
   moddat_g_str_DesGrp = cmb_ClaCre.Text
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_TIPPRV WHERE "
   g_str_Parame = g_str_Parame & "TIPPRV_TIPPRV = '" & moddat_g_str_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "TIPPRV_CLACRE = '" & moddat_g_str_CodGrp & "' "
   g_str_Parame = g_str_Parame & "ORDER BY TIPPRV_CLFCRE ASC, TIPPRV_CLAGAR ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = moddat_gf_ConsultaClasifCred(moddat_g_str_CodGrp, Trim(g_rst_Princi!TIPPRV_CLFCRE))
      
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_ConsultaClaseGar(Trim(g_rst_Princi!TipPrv_ClaGar))
      
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!TipPrv_Porcen, "###,##0.00")
      
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!TIPPRV_CLFCRE)
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!TipPrv_ClaGar)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub


