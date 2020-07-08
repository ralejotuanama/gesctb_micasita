VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mat_Produc_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8835
   ClientLeft      =   2220
   ClientTop       =   2805
   ClientWidth     =   15060
   Icon            =   "GesCtb_frm_172.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8805
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15060
      _Version        =   65536
      _ExtentX        =   26564
      _ExtentY        =   15531
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
         Height          =   6465
         Left            =   30
         TabIndex        =   9
         Top             =   2280
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
         BevelOuter      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   6045
            Left            =   30
            TabIndex        =   2
            Top             =   360
            Width           =   14900
            _ExtentX        =   26273
            _ExtentY        =   10663
            _Version        =   393216
            Rows            =   25
            Cols            =   15
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_ConCtb 
            Height          =   285
            Left            =   9630
            TabIndex        =   10
            Top             =   60
            Width           =   3070
            _Version        =   65536
            _ExtentX        =   5415
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   285
            Left            =   8550
            TabIndex        =   11
            Top             =   60
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
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
         Begin Threed.SSPanel pnl_CtaCtb 
            Height          =   285
            Left            =   12700
            TabIndex        =   12
            Top             =   60
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
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
         Begin Threed.SSPanel pnl_TipCre 
            Height          =   285
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Crédito"
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
         Begin Threed.SSPanel pnl_SitCre 
            Height          =   285
            Left            =   1760
            TabIndex        =   21
            Top             =   60
            Width           =   1690
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación Crédito"
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
         Begin Threed.SSPanel pnl_ClaGar 
            Height          =   285
            Left            =   3450
            TabIndex        =   22
            Top             =   60
            Width           =   1690
            _Version        =   65536
            _ExtentX        =   2999
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
         Begin Threed.SSPanel pnl_EmpSeg 
            Height          =   285
            Left            =   5150
            TabIndex        =   23
            Top             =   60
            Width           =   1690
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Emp. Seguros"
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
         Begin Threed.SSPanel pnl_GasCie 
            Height          =   285
            Left            =   6840
            TabIndex        =   24
            Top             =   60
            Width           =   1710
            _Version        =   65536
            _ExtentX        =   3016
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Gasto Cierre"
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
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
            Caption         =   "Cuentas Contables por Producto"
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
            Picture         =   "GesCtb_frm_172.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   16
         Top             =   780
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
            Picture         =   "GesCtb_frm_172.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_172.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_172.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14325
            Picture         =   "GesCtb_frm_172.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_172.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_172.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   18
         Top             =   1470
         Width           =   15000
         _Version        =   65536
         _ExtentX        =   26458
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
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   13800
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   13800
         End
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   20
            Top             =   390
            Width           =   825
         End
         Begin VB.Label Label4 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "frm_Mat_Produc_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_Empres()      As moddat_tpo_Genera

Private Sub cmb_Empres_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmb_Produc_Click()
   Call gs_SetFocus(cmb_Empres)
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_FlgAct = 1
   
   frm_Mat_Produc_02.Show 1
   
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
   
   grd_Listad.Col = 8
   moddat_g_str_TipCre = grd_Listad.Text

   grd_Listad.Col = 9
   moddat_g_str_SitCre = grd_Listad.Text

   grd_Listad.Col = 10
   moddat_g_str_ClaGar = grd_Listad.Text

   grd_Listad.Col = 11
   moddat_g_str_CodGrp = grd_Listad.Text

   grd_Listad.Col = 12
   moddat_g_str_CodIte = grd_Listad.Text

   grd_Listad.Col = 13
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   grd_Listad.Col = 14
   moddat_g_str_Codigo = Trim(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM CTB_CTAPRD "
   g_str_Parame = g_str_Parame & " WHERE CTAPRD_CODPRD = '" & moddat_g_str_CodPrd & "' "
   g_str_Parame = g_str_Parame & "   AND CTAPRD_TIPCRE = '" & moddat_g_str_TipCre & "' "
   g_str_Parame = g_str_Parame & "   AND CTAPRD_SITCRE = '" & moddat_g_str_SitCre & "' "
   g_str_Parame = g_str_Parame & "   AND CTAPRD_CLAGAR = '" & moddat_g_str_ClaGar & "' "
   g_str_Parame = g_str_Parame & "   AND CTAPRD_EMPSEG = '" & moddat_g_str_CodGrp & "' "
   g_str_Parame = g_str_Parame & "   AND CTAPRD_GASCIE = " & moddat_g_str_CodIte & " "
   g_str_Parame = g_str_Parame & "   AND CTAPRD_TIPMON = " & CStr(moddat_g_int_TipMon) & " "
   g_str_Parame = g_str_Parame & "   AND CTAPRD_CONCTB = '" & moddat_g_str_Codigo & "' "
   g_str_Parame = g_str_Parame & "   AND CTAPRD_EMPGRP = '" & moddat_g_str_CodEmp & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_Produc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Produc)
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

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 8
   moddat_g_str_TipCre = grd_Listad.Text

   grd_Listad.Col = 9
   moddat_g_str_SitCre = grd_Listad.Text

   grd_Listad.Col = 10
   moddat_g_str_ClaGar = grd_Listad.Text

   grd_Listad.Col = 11
   moddat_g_str_CodGrp = grd_Listad.Text

   grd_Listad.Col = 12
   moddat_g_str_CodIte = grd_Listad.Text

   grd_Listad.Col = 13
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   grd_Listad.Col = 14
   moddat_g_str_Codigo = Trim(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   
   frm_Mat_Produc_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_Produc)
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
   Call moddat_gs_Carga_Produc(cmb_Produc, l_arr_Produc, 4)
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)

   grd_Listad.ColWidth(0) = 1700
   grd_Listad.ColWidth(1) = 1700
   grd_Listad.ColWidth(2) = 1700
   grd_Listad.ColWidth(3) = 1700
   grd_Listad.ColWidth(4) = 1700
   grd_Listad.ColWidth(5) = 1100
   grd_Listad.ColWidth(6) = 3045
   grd_Listad.ColWidth(7) = 1700
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 0
   grd_Listad.ColWidth(14) = 0
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   cmb_Produc.ListIndex = -1
   cmb_Empres.ListIndex = -1
   
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Produc.Enabled = p_Activa
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
   
   moddat_g_str_CodPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo
   moddat_g_str_NomPrd = cmb_Produc.Text
   moddat_g_str_CodEmp = l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo
   moddat_g_str_RazSoc = cmb_Empres.Text
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   'g_str_Parame = g_str_Parame & "SELECT * FROM CTB_CTAPRD "
   'g_str_Parame = g_str_Parame & " WHERE CTAPRD_CODPRD = '" & moddat_g_str_CodPrd & "' "
   'g_str_Parame = g_str_Parame & "   AND CTAPRD_EMPGRP = '" & moddat_g_str_CodEmp & "' "
   'g_str_Parame = g_str_Parame & " ORDER BY CTAPRD_TIPMON ASC, CTAPRD_CONCTB ASC, CTAPRD_TIPCRE ASC, CTAPRD_SITCRE ASC, CTAPRD_CLAGAR ASC, CTAPRD_EMPSEG ASC, CTAPRD_GASCIE ASC"

   g_str_Parame = g_str_Parame & " SELECT A.*, TRIM(B.CONCTB_DESCRI) AS CONCEPTO_CTB  "
   g_str_Parame = g_str_Parame & "   FROM CTB_CTAPRD A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CTB_CONCTB B ON A.CTAPRD_CONCTB = B.CONCTB_CODCAM  "
   g_str_Parame = g_str_Parame & "  WHERE A.CTAPRD_CODPRD = '" & moddat_g_str_CodPrd & "' "
   g_str_Parame = g_str_Parame & "    AND A.CTAPRD_EMPGRP = '" & moddat_g_str_CodEmp & "' "
   g_str_Parame = g_str_Parame & "  ORDER BY A.CTAPRD_TIPMON ASC, CONCEPTO_CTB ASC, A.CTAPRD_TIPCRE ASC,  "
   g_str_Parame = g_str_Parame & "          A.CTAPRD_SITCRE ASC, A.CTAPRD_CLAGAR ASC, A.CTAPRD_EMPSEG ASC, A.CTAPRD_GASCIE ASC  "
   
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
      If Trim(g_rst_Princi!CTAPRD_TIPCRE) <> "999" Then
         grd_Listad.Text = moddat_gf_Consulta_TipoCreditoCtb(Trim(g_rst_Princi!CTAPRD_TIPCRE))
      Else
         grd_Listad.Text = "<< NO APLICA >>"
      End If
      
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(g_rst_Princi!CTAPRD_TIPCRE)
      
      grd_Listad.Col = 1
      If Trim(g_rst_Princi!CtaPrd_SitCre) <> "999" Then
         grd_Listad.Text = moddat_gf_Consulta_SituacionCreditoCtb("4", Trim(g_rst_Princi!CtaPrd_SitCre))
      Else
         grd_Listad.Text = "<< NO APLICA >>"
      End If
      
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(g_rst_Princi!CtaPrd_SitCre)
      
      grd_Listad.Col = 2
      If Trim(g_rst_Princi!CTAPRD_CLAGAR) <> "999" Then
         grd_Listad.Text = moddat_gf_ConsultaClaseGar(Trim(g_rst_Princi!CTAPRD_CLAGAR))
      Else
         grd_Listad.Text = "<< NO APLICA >>"
      End If
      
      grd_Listad.Col = 10
      grd_Listad.Text = Trim(g_rst_Princi!CTAPRD_CLAGAR)
      
      grd_Listad.Col = 3
      If Trim(g_rst_Princi!CTAPRD_EMPSEG) <> "999999" Then
         grd_Listad.Text = moddat_gf_Consulta_ComSeg(Trim(g_rst_Princi!CTAPRD_EMPSEG))
      Else
         grd_Listad.Text = "<< NO APLICA >>"
      End If
      
      grd_Listad.Col = 11
      grd_Listad.Text = Trim(g_rst_Princi!CTAPRD_EMPSEG)
      
      grd_Listad.Col = 4
      If g_rst_Princi!CTAPRD_GASCIE <> 99 Then
         grd_Listad.Text = moddat_gf_Consulta_ParDes("265", CStr(g_rst_Princi!CTAPRD_GASCIE))
      Else
         grd_Listad.Text = "<< NO APLICA >>"
      End If
      
      grd_Listad.Col = 12
      grd_Listad.Text = CStr(g_rst_Princi!CTAPRD_GASCIE)
      
      grd_Listad.Col = 5
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!CTAPRD_TIPMON))
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!CONCEPTO_CTB & "")
      'grd_Listad.Text = moddat_gf_Consulta_ConceptoCtb(g_rst_Princi!CTAPRD_CONCTB)
      
      grd_Listad.Col = 7
      grd_Listad.Text = Trim(g_rst_Princi!CtaPrd_CtaCtb)
      
      grd_Listad.Col = 13
      grd_Listad.Text = CStr(g_rst_Princi!CTAPRD_TIPMON)
      
      grd_Listad.Col = 14
      grd_Listad.Text = Trim(g_rst_Princi!CTAPRD_CONCTB) 'CONCEPTO CONTABLE
      
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

Private Sub pnl_ClaGar_Click()
   If pnl_ClaGar.Tag = "" Then
      pnl_ClaGar.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_ClaGar.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_ConCtb_Click()
   If pnl_ConCtb.Tag = "" Then
      pnl_ConCtb.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_ConCtb.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

Private Sub pnl_CtaCtb_Click()
   If pnl_CtaCtb.Tag = "" Then
      pnl_CtaCtb.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 7, "N")
   Else
      pnl_CtaCtb.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 7, "N-")
   End If
End Sub

Private Sub pnl_EmpSeg_Click()
   If pnl_EmpSeg.Tag = "" Then
      pnl_EmpSeg.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_EmpSeg.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_GasCie_Click()
   If pnl_GasCie.Tag = "" Then
      pnl_GasCie.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_GasCie.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Moneda_Click()
   If pnl_Moneda.Tag = "" Then
      pnl_Moneda.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Moneda.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_SitCre_Click()
   If pnl_SitCre.Tag = "" Then
      pnl_SitCre.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_SitCre.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_TipCre_Click()
   If pnl_TipCre.Tag = "" Then
      pnl_TipCre.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_TipCre.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub


