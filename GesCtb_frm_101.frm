VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_Period_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   11160
   ClientTop       =   3465
   ClientWidth     =   7170
   Icon            =   "GesCtb_frm_101.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8115
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
      _ExtentY        =   14314
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
         Height          =   675
         Left            =   30
         TabIndex        =   9
         Top             =   750
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
         Begin VB.CommandButton cmd_CerPer 
            Height          =   585
            Left            =   3030
            Picture         =   "GesCtb_frm_101.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Cerrar Período"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_101.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_101.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_101.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6450
            Picture         =   "GesCtb_frm_101.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_101.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_101.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   5775
         Left            =   30
         TabIndex        =   10
         Top             =   2280
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   10186
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
            Height          =   5385
            Left            =   30
            TabIndex        =   7
            Top             =   360
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   9499
            _Version        =   393216
            Rows            =   18
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_CodEmp 
            Height          =   285
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Año"
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
         Begin Threed.SSPanel pnl_Tit_RazSoc 
            Height          =   285
            Left            =   1050
            TabIndex        =   12
            Top             =   60
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mes"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   2970
            TabIndex        =   17
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
            Left            =   4830
            TabIndex        =   18
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Bloqueo Registros"
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
            Top             =   90
            Width           =   3915
            _Version        =   65536
            _ExtentX        =   6906
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de Períodos"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   6630
            Top             =   30
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   90
            Picture         =   "GesCtb_frm_101.frx":168A
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   765
         Left            =   30
         TabIndex        =   15
         Top             =   1470
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
         Begin VB.ComboBox cmb_TipPer 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   390
            Width           =   5895
         End
         Begin VB.ComboBox cmb_EmpGrp 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5895
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Período:"
            Height          =   225
            Left            =   60
            TabIndex        =   20
            Top             =   390
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   225
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_Period_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_EmpGrp()   As moddat_tpo_Genera

Private Sub cmb_EmpGrp_Click()
   Call gs_SetFocus(cmb_TipPer)
End Sub

Private Sub cmb_EmpGrp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpGrp_Click
   End If
End Sub

Private Sub cmb_TipPer_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_TipPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPer_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_FlgAct = 1
   
   frm_Mnt_Period_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Borrar_Click()
   Dim r_str_Situac     As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   
   grd_Listad.Col = 4
   moddat_g_str_CodIte = grd_Listad.Text
   
   grd_Listad.Col = 5
   r_str_Situac = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)

   If CInt(r_str_Situac) = 9 Then
      MsgBox "No se puede eliminar el registro porque el Período ya fue Cerrado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = "DELETE FROM CTB_PERMES WHERE "
   g_str_Parame = g_str_Parame & "PERMES_CODEMP = '" & moddat_g_str_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PERMES_TIPPER = " & moddat_g_str_CodMod & " AND "
   g_str_Parame = g_str_Parame & "PERMES_CODANO = " & moddat_g_str_Codigo & " AND "
   g_str_Parame = g_str_Parame & "PERMES_CODMES = " & moddat_g_str_CodIte & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_EmpGrp.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa del Grupo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EmpGrp)
      Exit Sub
   End If
   
   If cmb_TipPer.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Período.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipPer)
      Exit Sub
   End If
   
   moddat_g_str_CodGrp = l_arr_EmpGrp(cmb_EmpGrp.ListIndex + 1).Genera_Codigo
   moddat_g_str_DesGrp = cmb_EmpGrp.Text
   
   moddat_g_str_CodMod = CStr(cmb_TipPer.ItemData(cmb_TipPer.ListIndex))
   moddat_g_str_DesMod = cmb_TipPer.Text
   
   Call fs_Activa(False)
   Call fs_Buscar
End Sub

Private Sub cmd_CerPer_Click()
   Dim r_str_Situac     As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   
   grd_Listad.Col = 4
   moddat_g_str_CodIte = grd_Listad.Text
   
   grd_Listad.Col = 5
   r_str_Situac = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)

   If CInt(r_str_Situac) = 9 Then
      MsgBox "No se puede cerrar el Período ya fue Cerrado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   If MsgBox("¿Está seguro que desea cerrar el Período?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CTB_PERMES_CERRAR ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
      g_str_Parame = g_str_Parame & moddat_g_str_CodMod & ", "
      g_str_Parame = g_str_Parame & moddat_g_str_Codigo & ", "
      g_str_Parame = g_str_Parame & moddat_g_str_CodIte & ", "
      g_str_Parame = g_str_Parame & "9, "
         
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CTB_PERMES. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop

   Call fs_Buscar

   Screen.MousePointer = 0
End Sub

Private Sub cmd_Editar_Click()
   Dim r_str_Situac     As String
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   
   grd_Listad.Col = 4
   moddat_g_str_CodIte = grd_Listad.Text
   
   grd_Listad.Col = 5
   r_str_Situac = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If CInt(r_str_Situac) = 9 Then
      MsgBox "No se puede modificar los datos porque el Período ya fue cerrado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   
   frm_Mnt_Period_02.Show 1
   
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
   grd_Listad.ColWidth(0) = 1005
   grd_Listad.ColWidth(1) = 1935
   grd_Listad.ColWidth(2) = 1875
   grd_Listad.ColWidth(3) = 1875
   grd_Listad.ColWidth(4) = 0
   grd_Listad.ColWidth(5) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter

   Call moddat_gs_Carga_EmpGrp(cmb_EmpGrp, l_arr_EmpGrp)
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPer, 1, "251")
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   cmd_CerPer.Enabled = False
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = "SELECT * FROM CTB_PERMES WHERE "
   g_str_Parame = g_str_Parame & "PERMES_CODEMP = '" & moddat_g_str_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "PERMES_TIPPER = " & moddat_g_str_CodMod & " "
   g_str_Parame = g_str_Parame & "ORDER BY PERMES_CODANO DESC, PERMES_CODMES DESC "
   
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
      grd_Listad.Text = Format(g_rst_Princi!PERMES_CODANO, "0000")
      
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_ParDes("033", CStr(g_rst_Princi!PERMES_CODMES))
      
      grd_Listad.Col = 2
      grd_Listad.Text = moddat_gf_Consulta_ParDes("250", CStr(g_rst_Princi!PERMES_SITUAC))
      
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!PERMES_BLQREG))
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!PERMES_CODMES)
      
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!PERMES_SITUAC)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      cmd_CerPer.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Limpia()
   cmb_EmpGrp.ListIndex = -1
   cmb_TipPer.ListIndex = -1
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_EmpGrp.Enabled = p_Activa
   cmb_TipPer.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
   cmd_CerPer.Enabled = Not p_Activa
   grd_Listad.Enabled = Not p_Activa
End Sub

