VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Mnt_EFBG_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "GesCtb_frm_908.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6270
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   11060
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
         TabIndex        =   15
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
            TabIndex        =   16
            Top             =   60
            Width           =   6435
            _Version        =   65536
            _ExtentX        =   11351
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de EEFF - Balance General"
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
            Picture         =   "GesCtb_frm_908.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3045
         Left            =   30
         TabIndex        =   17
         Top             =   1950
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   5371
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
            Height          =   2655
            Left            =   30
            TabIndex        =   10
            Top             =   360
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   4683
            _Version        =   393216
            Rows            =   25
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   840
            TabIndex        =   18
            Top             =   60
            Width           =   6195
            _Version        =   65536
            _ExtentX        =   10927
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción"
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
            TabIndex        =   19
            Top             =   60
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código"
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
         TabIndex        =   20
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
         Begin VB.CommandButton cmd_DetGrp 
            Height          =   585
            Left            =   3015
            Picture         =   "GesCtb_frm_908.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Detalle de SubGrupo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_908.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2400
            Picture         =   "GesCtb_frm_908.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1800
            Picture         =   "GesCtb_frm_908.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1200
            Picture         =   "GesCtb_frm_908.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6750
            Picture         =   "GesCtb_frm_908.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_908.frx":17C2
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   3630
            Picture         =   "GesCtb_frm_908.frx":1ACC
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   4230
            Picture         =   "GesCtb_frm_908.frx":1F0E
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Cancelar "
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   435
         Left            =   30
         TabIndex        =   21
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
         Begin VB.ComboBox cmb_TipEEFF 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5655
         End
         Begin VB.Label Label4 
            Caption         =   "Grupo EEFF:"
            Height          =   255
            Left            =   90
            TabIndex        =   22
            Top             =   90
            Width           =   1665
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   1155
         Left            =   30
         TabIndex        =   23
         Top             =   5055
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   2037
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
         Begin VB.TextBox txt_TipCta 
            Height          =   315
            Left            =   1860
            MaxLength       =   1
            TabIndex        =   13
            Top             =   750
            Width           =   555
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1860
            MaxLength       =   3
            TabIndex        =   11
            Top             =   90
            Width           =   555
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   12
            Top             =   420
            Width           =   5445
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Tipo de Cuenta"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   26
            Top             =   780
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Código Clasificación:"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   25
            Top             =   120
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Descripción:"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   24
            Top             =   450
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_EFBG_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_TipEEFF()      As moddat_tpo_Genera

Private Sub cmd_Buscar_Click()
   If cmb_TipEEFF.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo EEFF.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipEEFF)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Activa2(True)
   Call fs_Activa(False)
   Call fs_Buscar
   Call fs_Limpia
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   cmb_TipEEFF.ListIndex = -1
   cmd_DetGrp.Enabled = False
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Activa2(True)
   Call fs_Activa(True)
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipEEFF)
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_FlgAct = 1
   
   Screen.MousePointer = 11
   Call fs_Activa2(False)
   Call fs_Activa(True)
   cmb_TipEEFF.Enabled = False
   cmd_Buscar.Enabled = False
   Call fs_Limpia
   Call gs_SetFocus(txt_Codigo)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_int_FlgGrb = 2
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   moddat_g_str_CodIte = grd_Listad.Text
   Screen.MousePointer = 11
   Call fs_Activa2(False)
   Call fs_Activa(True)
   cmb_TipEEFF.Enabled = False
   cmd_Buscar.Enabled = False
   Call fs_Limpia
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DISTINCT CODG_SBGRP, NOMB_SBGRP, TIPO_CTACTB "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_EEBG "
   g_str_Parame = g_str_Parame & " WHERE TRIM(INDC_TIPO) IN ('D','S') "
   g_str_Parame = g_str_Parame & "   AND CODG_GRUPO = " & moddat_g_str_Codigo & "  "
   g_str_Parame = g_str_Parame & "   AND CODG_SBGRP = " & CInt(moddat_g_str_CodIte) & "  "
   g_str_Parame = g_str_Parame & "ORDER BY CODG_SBGRP "
   
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
   txt_Codigo.Text = g_rst_Princi!CODG_SBGRP
   txt_Descri.Text = Trim(g_rst_Princi!NOMB_SBGRP)
   txt_TipCta.Text = Trim(g_rst_Princi!TIPO_CTACTB)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   txt_Codigo.Enabled = False
   Call gs_SetFocus(txt_Descri)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_CodIte = grd_Listad.Text
   Call gs_RefrescaGrid(grd_Listad)
   
   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = "USP_CTB_EEFF_BG ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_Descri & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & 3 & ","
   g_str_Parame = g_str_Parame & 1 & ")"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_DetGrp_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_CodIte = Trim(grd_Listad.Text)
   grd_Listad.Col = 1
   moddat_g_str_DesIte = Trim(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   
   frm_Mnt_EFBG_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_Codigo.Text)) = 0 Then
      MsgBox "Debe ingresar el Código de Clasificación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Codigo)
      Exit Sub
   End If
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If
   If Len(Trim(txt_TipCta.Text)) = 0 Then
      MsgBox "Debe ingresar el tipo de Cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_TipCta)
      Exit Sub
   End If
   If Not (Trim(txt_TipCta.Text) <> "H" Or Trim(txt_TipCta.Text) <> "D") Then
      MsgBox "El tipo de cuenta solo puede ser 'D'-Debe o 'H'-Haber.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_TipCta)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   If moddat_g_int_FlgGrb = 1 Then
      
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CTB_EEFF_BG ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_Descri & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Codigo.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Descri.Text) & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_TipCta.Text) & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & 1 & ","
         g_str_Parame = g_str_Parame & 1 & ")"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_FlgGOK = True
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   Else
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CTB_EEFF_BG ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_Descri & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Descri.Text) & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_TipCta.Text) & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & 2 & ","
         g_str_Parame = g_str_Parame & 1 & ")"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_FlgGOK = True
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   End If
   
   Call fs_Buscar
   Call fs_Activa2(True)
   Call fs_Activa(False)
   Call fs_Limpia
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Activa2(True)
   Call fs_Activa(False)
   Call fs_Limpia
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_TipEEFF)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 765
   grd_Listad.ColWidth(1) = 6195
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter

   cmd_DetGrp.Enabled = False
   Call moddat_gs_Carga_TipBalGen(cmb_TipEEFF, l_arr_TipEEFF)
End Sub

Private Sub fs_Limpia()
   txt_Codigo.Text = ""
   txt_Descri.Text = ""
   txt_TipCta.Text = ""
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_TipEEFF.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
End Sub

Private Sub fs_Activa2(ByVal p_Activa As Integer)
   cmd_Agrega.Enabled = p_Activa
   cmd_Editar.Enabled = p_Activa
   cmd_Borrar.Enabled = p_Activa
   cmd_Grabar.Enabled = Not p_Activa
   cmd_Cancel.Enabled = Not p_Activa
   txt_Codigo.Enabled = Not p_Activa
   txt_Descri.Enabled = Not p_Activa
   txt_TipCta.Enabled = Not p_Activa
   SSPanel6.Enabled = Not p_Activa
End Sub

Public Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   moddat_g_str_Codigo = l_arr_TipEEFF(cmb_TipEEFF.ListIndex + 1).Genera_Codigo
   moddat_g_str_Descri = cmb_TipEEFF.Text
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DISTINCT CODG_SBGRP, NOMB_SBGRP "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_EEBG "
   g_str_Parame = g_str_Parame & " WHERE TRIM(INDC_TIPO) IN ('D','S') "
   g_str_Parame = g_str_Parame & "   AND CODG_GRUPO = " & moddat_g_str_Codigo & " "
   g_str_Parame = g_str_Parame & "   AND NOMB_SBGRP IS NOT NULL "
   g_str_Parame = g_str_Parame & " ORDER BY CODG_SBGRP"
   
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
      grd_Listad.Text = g_rst_Princi!CODG_SBGRP
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!NOMB_SBGRP)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Editar.Enabled = True
      cmd_Borrar.Enabled = True
      grd_Listad.Enabled = True
      cmd_DetGrp.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_DetGrp_Click
End Sub

Private Sub moddat_gs_Carga_TipBalGen(p_Combo As ComboBox, p_Arregl() As moddat_tpo_Genera)
   p_Combo.Clear
   ReDim p_Arregl(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT CODG_GRUPO, NOMB_GRUPO "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_EEBG "
   g_str_Parame = g_str_Parame & " WHERE TRIM(INDC_TIPO) IN ('G','T') "
   g_str_Parame = g_str_Parame & "ORDER BY CODG_GRUPO "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      p_Combo.AddItem Trim(g_rst_Listas!NOMB_GRUPO)
      ReDim Preserve p_Arregl(UBound(p_Arregl) + 1)
      
      p_Arregl(UBound(p_Arregl)).Genera_Codigo = Trim(g_rst_Listas!CODG_GRUPO)
      p_Arregl(UBound(p_Arregl)).Genera_Nombre = Trim(g_rst_Listas!NOMB_GRUPO)
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
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
      Call gs_SetFocus(txt_TipCta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & " ,-/\[]().")
   End If
End Sub

Private Sub txt_TipCta_GotFocus()
   Call gs_SelecTodo(txt_TipCta)
End Sub

Private Sub txt_TipCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, "HD")
   End If
End Sub
