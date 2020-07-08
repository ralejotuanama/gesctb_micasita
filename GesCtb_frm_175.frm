VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Mat_CtaPry_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6495
   ClientLeft      =   3045
   ClientTop       =   2220
   ClientWidth     =   11340
   Icon            =   "GesCtb_frm_175.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6465
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11325
      _Version        =   65536
      _ExtentX        =   19976
      _ExtentY        =   11404
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
         Height          =   4125
         Left            =   30
         TabIndex        =   9
         Top             =   2280
         Width           =   11235
         _Version        =   65536
         _ExtentX        =   19817
         _ExtentY        =   7276
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
            Height          =   3735
            Left            =   30
            TabIndex        =   10
            Top             =   360
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   6588
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
         Begin Threed.SSPanel pnl_Tit_Descri 
            Height          =   285
            Left            =   3330
            TabIndex        =   11
            Top             =   60
            Width           =   4305
            _Version        =   65536
            _ExtentX        =   7594
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Concepto Contable"
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
         Begin Threed.SSPanel pnl_Tit_Codigo 
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   3285
            _Version        =   65536
            _ExtentX        =   5794
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
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
            Left            =   7620
            TabIndex        =   13
            Top             =   60
            Width           =   3225
            _Version        =   65536
            _ExtentX        =   5689
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuenta Contable"
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
         TabIndex        =   14
         Top             =   60
         Width           =   11235
         _Version        =   65536
         _ExtentX        =   19817
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
            TabIndex        =   15
            Top             =   60
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Cuentas Contables por Proyecto Hipotecario"
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
            Picture         =   "GesCtb_frm_175.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   16
         Top             =   780
         Width           =   11235
         _Version        =   65536
         _ExtentX        =   19817
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
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_175.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_175.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_175.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10590
            Picture         =   "GesCtb_frm_175.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_175.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_175.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   17
         Top             =   1470
         Width           =   11235
         _Version        =   65536
         _ExtentX        =   19817
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
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   9375
         End
         Begin VB.ComboBox cmb_PryVin 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   9375
         End
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   390
            Width           =   1605
         End
         Begin VB.Label Label4 
            Caption         =   "Proyecto Hipotecario:"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1605
         End
      End
   End
End
Attribute VB_Name = "frm_Mat_CtaPry_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_arr_PryVin()      As moddat_tpo_Genera

Private Sub cmb_Empres_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmb_PryVin_Click()
   Call gs_SetFocus(cmb_Empres)
End Sub

Private Sub cmb_PryVin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PryVin_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_FlgAct = 1
   
   frm_Mat_CtaPry_02.Show 1
   
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
   moddat_g_int_TipMon = CStr(grd_Listad.Text)
   
   grd_Listad.Col = 4
   moddat_g_str_Codigo = Trim(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = "DELETE FROM CTB_CTAPRY WHERE "
   g_str_Parame = g_str_Parame & "CTAPRY_CODPRY = '" & moddat_g_str_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "CTAPRY_TIPMON = " & CStr(moddat_g_int_TipMon) & " AND "
   g_str_Parame = g_str_Parame & "CTAPRY_CONCTB = '" & moddat_g_str_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "CTAPRY_EMPGRP = '" & moddat_g_str_CodEmp & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0

End Sub

Private Sub cmd_Buscar_Click()
   If cmb_PryVin.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Proyecto Hipotecario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PryVin)
      Exit Sub
   End If

   If cmb_Empres.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Empres)
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
   moddat_g_int_TipMon = CStr(grd_Listad.Text)
   
   grd_Listad.Col = 4
   moddat_g_str_Codigo = Trim(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   
   frm_Mat_CtaPry_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_PryVin)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call gs_CentraForm(Me)
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_PryVin(cmb_PryVin, l_arr_PryVin)
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)

   grd_Listad.ColWidth(0) = 3275
   grd_Listad.ColWidth(1) = 4305
   grd_Listad.ColWidth(2) = 3215
   grd_Listad.ColWidth(3) = 0
   grd_Listad.ColWidth(4) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   cmb_PryVin.ListIndex = -1
   cmb_Empres.ListIndex = -1
   
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_PryVin.Enabled = p_Activa
   cmb_Empres.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   moddat_g_str_CodGrp = l_arr_PryVin(cmb_PryVin.ListIndex + 1).Genera_Codigo
   moddat_g_str_DesGrp = cmb_PryVin.Text
   
   moddat_g_str_CodEmp = l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo
   moddat_g_str_RazSoc = cmb_Empres.Text
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAPRY WHERE "
   g_str_Parame = g_str_Parame & "CTAPRY_CODPRY = '" & moddat_g_str_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "CTAPRY_EMPGRP = '" & moddat_g_str_CodEmp & "' "
   g_str_Parame = g_str_Parame & "ORDER BY CTAPRY_TIPMON ASC, CTAPRY_CONCTB ASC"

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
      grd_Listad.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!CTAPRY_TIPMON))
      
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_ConceptoCtb(g_rst_Princi!CTAPRY_CONCTB)
      
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(g_rst_Princi!CTAPRY_CTACTB)
      
      grd_Listad.Col = 3
      grd_Listad.Text = CStr(g_rst_Princi!CTAPRY_TIPMON)
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!CTAPRY_CONCTB)
      
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

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub


