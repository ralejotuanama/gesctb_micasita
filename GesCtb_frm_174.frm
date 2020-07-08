VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Mat_MatCon_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9180
   ClientLeft      =   5355
   ClientTop       =   2820
   ClientWidth     =   15060
   Icon            =   "GesCtb_frm_174.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9165
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   15075
      _Version        =   65536
      _ExtentX        =   26591
      _ExtentY        =   16166
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
         Left            =   60
         TabIndex        =   21
         Top             =   60
         Width           =   14955
         _Version        =   65536
         _ExtentX        =   26379
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
            TabIndex        =   22
            Top             =   60
            Width           =   5955
            _Version        =   65536
            _ExtentX        =   10504
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Dinámicas Contables"
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
            Picture         =   "GesCtb_frm_174.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   23
         Top             =   780
         Width           =   14955
         _Version        =   65536
         _ExtentX        =   26379
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
         Begin VB.CommandButton cmd_Acepta 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_174.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Grabar Detalle"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   3060
            Picture         =   "GesCtb_frm_174.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Cancelar Edición Detalle"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_174.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1260
            Picture         =   "GesCtb_frm_174.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_174.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14340
            Picture         =   "GesCtb_frm_174.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   60
            Picture         =   "GesCtb_frm_174.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   435
         Left            =   60
         TabIndex        =   24
         Top             =   1470
         Width           =   14955
         _Version        =   65536
         _ExtentX        =   26379
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
         Begin Threed.SSPanel pnl_Empres 
            Height          =   315
            Left            =   1590
            TabIndex        =   25
            Top             =   60
            Width           =   13305
            _Version        =   65536
            _ExtentX        =   23469
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
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1305
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3555
         Left            =   60
         TabIndex        =   27
         Top             =   3090
         Width           =   14955
         _Version        =   65536
         _ExtentX        =   26379
         _ExtentY        =   6271
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
            Height          =   3165
            Left            =   30
            TabIndex        =   5
            Top             =   360
            Width           =   14895
            _ExtentX        =   26273
            _ExtentY        =   5583
            _Version        =   393216
            Rows            =   25
            Cols            =   12
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_Descri 
            Height          =   285
            Left            =   2220
            TabIndex        =   28
            Top             =   60
            Width           =   1845
            _Version        =   65536
            _ExtentX        =   3254
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Concepto"
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
            Left            =   9060
            TabIndex        =   29
            Top             =   60
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "D/H"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   10020
            TabIndex        =   30
            Top             =   60
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Cambio"
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
         Begin Threed.SSPanel SSPanel6 
            Height          =   285
            Left            =   11760
            TabIndex        =   31
            Top             =   60
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Concepto Operativo"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   4050
            TabIndex        =   36
            Top             =   60
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   7080
            TabIndex        =   37
            Top             =   60
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   60
            TabIndex        =   48
            Top             =   60
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
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
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   1095
         Left            =   60
         TabIndex        =   32
         Top             =   1950
         Width           =   14955
         _Version        =   65536
         _ExtentX        =   26379
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
         Begin VB.ComboBox cmb_LibCtb 
            Height          =   315
            Left            =   9690
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   720
            Width           =   5205
         End
         Begin VB.ComboBox cmb_TipMat 
            Height          =   315
            Left            =   9690
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   390
            Width           =   5205
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1590
            MaxLength       =   10
            TabIndex        =   0
            Text            =   "0"
            Top             =   60
            Width           =   1425
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1590
            MaxLength       =   250
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   5715
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   5715
         End
         Begin VB.Label Label5 
            Caption         =   "Libro Contable:"
            Height          =   285
            Left            =   8160
            TabIndex        =   39
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Matriz:"
            Height          =   285
            Left            =   8160
            TabIndex        =   38
            Top             =   390
            Width           =   885
         End
         Begin VB.Label Label8 
            Caption         =   "Código Matriz:"
            Height          =   285
            Left            =   60
            TabIndex        =   35
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label2 
            Caption         =   "Descripción:"
            Height          =   285
            Left            =   60
            TabIndex        =   34
            Top             =   390
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda:"
            Height          =   285
            Left            =   60
            TabIndex        =   33
            Top             =   720
            Width           =   1425
         End
      End
      Begin Threed.SSPanel SSPanel13 
         Height          =   2415
         Left            =   60
         TabIndex        =   40
         Top             =   6690
         Width           =   14955
         _Version        =   65536
         _ExtentX        =   26379
         _ExtentY        =   4260
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
         Begin VB.ComboBox cmb_CtaCtb 
            Height          =   315
            Left            =   1590
            TabIndex        =   8
            Text            =   "cmb_CtaCtb"
            Top             =   720
            Width           =   13305
         End
         Begin VB.ComboBox cmb_TipCam 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2040
            Width           =   2175
         End
         Begin VB.ComboBox cmb_DebHab 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1710
            Width           =   2175
         End
         Begin VB.ComboBox cmb_ConOpe 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1380
            Width           =   13305
         End
         Begin VB.ComboBox cmb_ConCtb 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1050
            Width           =   13305
         End
         Begin VB.TextBox txt_DesIte 
            Height          =   315
            Left            =   1590
            MaxLength       =   250
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   60
            Width           =   13305
         End
         Begin VB.ComboBox cmb_TipCon 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   390
            Width           =   13305
         End
         Begin VB.Label Label13 
            Caption         =   "Tipo de Cambio:"
            Height          =   285
            Left            =   60
            TabIndex        =   47
            Top             =   2040
            Width           =   1185
         End
         Begin VB.Label Label12 
            Caption         =   "Debe / Haber:"
            Height          =   285
            Left            =   60
            TabIndex        =   46
            Top             =   1710
            Width           =   1185
         End
         Begin VB.Label Label11 
            Caption         =   "Concepto Operativo:"
            Height          =   285
            Left            =   60
            TabIndex        =   45
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label10 
            Caption         =   "Concepto Contable:"
            Height          =   285
            Left            =   60
            TabIndex        =   44
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label Label9 
            Caption         =   "Descripción:"
            Height          =   285
            Left            =   60
            TabIndex        =   43
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo Concepto:"
            Height          =   285
            Left            =   60
            TabIndex        =   42
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label6 
            Caption         =   "Cuenta Contable:"
            Height          =   285
            Left            =   60
            TabIndex        =   41
            Top             =   750
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_Mat_MatCon_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_TopNiv        As Integer
Dim l_int_FlgCmb        As Integer
Dim l_str_CtaCtb        As String
Dim l_int_FlgGrb        As Integer
Dim l_arr_ConCtb()      As moddat_tpo_Genera
Dim l_arr_CtaCtb()      As moddat_tpo_Genera
Dim l_arr_ParEmp()      As moddat_tpo_Genera
Dim l_arr_ConOpe()      As moddat_tpo_Genera
Dim l_arr_TipMat()      As moddat_tpo_Genera

Private Sub cmd_Grabar_Click()
Dim r_int_Contad     As Integer
   
   If Len(Trim(txt_Codigo.Text)) <> 10 Then
      MsgBox "Ingrese el Código de Matriz (El Código de Matriz debe tener 10 caracteres).", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Codigo)
      Exit Sub
   End If
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Ingrese la Descripción de la Matriz.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If
   If cmb_TipMat.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Matriz.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMat)
      Exit Sub
   End If
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   If cmb_LibCtb.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Libro Contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_LibCtb)
      Exit Sub
   End If
   If grd_Listad.Rows = 0 Then
      MsgBox "Debe ingresar los Detalles de la Matriz.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Agrega)
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 1 Then
      'Validar que el registro no exista
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM CTB_MATCAB "
      g_str_Parame = g_str_Parame & " WHERE MATCAB_CODMAT = '" & txt_Codigo.Text & "'"
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         MsgBox "El Código de Matriz ya ha sido registrado. Por favor verifique e intente nuevamente.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If moddat_g_int_FlgGrb = 2 Then
      'Borrando en Detalle
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "DELETE FROM CTB_MATDET "
      g_str_Parame = g_str_Parame & " WHERE MATDET_CODMAT = '" & txt_Codigo.Text & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   End If
   
   'Insertando Detalles
   grd_Listad.Redraw = False
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      g_str_Parame = "USP_CTB_MATDET ("
      g_str_Parame = g_str_Parame & "'" & Trim(txt_Codigo.Text) & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_Contad + 1) & ", "
      
      grd_Listad.Col = 0
      g_str_Parame = g_str_Parame & "'" & grd_Listad.Text & "', "
      
      'Tipo de Concepto
      grd_Listad.Col = 7
      g_str_Parame = g_str_Parame & grd_Listad.Text & ", "
         
      'Concepto Contable
      grd_Listad.Col = 8
      g_str_Parame = g_str_Parame & "'" & grd_Listad.Text & "', "
         
      'Cuenta Contable
      grd_Listad.Col = 3
      g_str_Parame = g_str_Parame & "'" & grd_Listad.Text & "', "
         
      'Tipo de Cambio
      grd_Listad.Col = 11
      g_str_Parame = g_str_Parame & grd_Listad.Text & ", "
         
      'Debe - Haber
      grd_Listad.Col = 10
      g_str_Parame = g_str_Parame & grd_Listad.Text & ", "
         
      'Concepto Operativo
      grd_Listad.Col = 9
      g_str_Parame = g_str_Parame & "'" & grd_Listad.Text & "', "
         
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "Error al ejecutar procedimiento USP_CTB_MATDET", vbCritical, modgen_g_str_NomPlt
      End If
   Next r_int_Contad
   grd_Listad.Redraw = True
   
   'Insertando Cabecera
   g_str_Parame = "USP_CTB_MATCAB ("
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Codigo.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Descri.Text) & "', "
   g_str_Parame = g_str_Parame & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "'" & l_arr_TipMat(cmb_TipMat.ListIndex + 1).Genera_Codigo & "', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodEmp & "', "
   g_str_Parame = g_str_Parame & CStr(cmb_LibCtb.ItemData(cmb_LibCtb.ListIndex)) & ", "
      
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ")"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "Error al ejecutar procedimiento USP_CTB_MATCAB", vbCritical, modgen_g_str_NomPlt
   End If
   
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Agrega_Click()
   l_int_FlgGrb = 1
   Call fs_LimpiaItem
   Call fs_Activa(True)
   Call gs_SetFocus(txt_DesIte)
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   l_int_FlgGrb = 2
   Call fs_Activa(True)
   
   'Pasando de grid a Campos
   grd_Listad.Col = 0
   txt_DesIte.Text = grd_Listad.Text
   
   grd_Listad.Col = 7
   Call gs_BuscarCombo_Item(cmb_TipCon, CInt(grd_Listad.Text))
   Call cmb_TipCon_Click
         
   If cmb_TipCon.ItemData(cmb_TipCon.ListIndex) = 1 Then
      grd_Listad.Col = 3
      cmb_CtaCtb.ListIndex = gf_Busca_Arregl(l_arr_CtaCtb, grd_Listad.Text) - 1
   Else
      grd_Listad.Col = 8
      cmb_ConCtb.ListIndex = gf_Busca_Arregl(l_arr_ConCtb, grd_Listad.Text) - 1
   End If
         
   grd_Listad.Col = 9
   cmb_ConOpe.ListIndex = gf_Busca_Arregl(l_arr_ConOpe, grd_Listad.Text) - 1
      
   grd_Listad.Col = 10
   Call gs_BuscarCombo_Item(cmb_DebHab, CInt(grd_Listad.Text))
      
   grd_Listad.Col = 11
   Call gs_BuscarCombo_Item(cmb_TipCam, CInt(grd_Listad.Text))
   
   Call gs_RefrescaGrid(grd_Listad)
   Call gs_SetFocus(txt_DesIte)
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Esta seguro de borrar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   If grd_Listad.Rows = 1 Then
      Call gs_LimpiaGrid(grd_Listad)
   Else
      grd_Listad.RemoveItem (grd_Listad.Row)
      Call gs_UbiIniGrid(grd_Listad)
   End If
End Sub

Private Sub cmd_Acepta_Click()
   If Len(Trim(txt_DesIte.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción del detalle.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DesIte)
      Exit Sub
   End If
   If cmb_TipCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Concepto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCon)
      Exit Sub
   End If
   If cmb_TipCon.ItemData(cmb_TipCon.ListIndex) = 1 Then
      If cmb_CtaCtb.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Cuenta Contable.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CtaCtb)
         Exit Sub
      End If
   Else
      If cmb_ConCtb.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Concepto Contable.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_ConCtb)
         Exit Sub
      End If
   End If
   If cmb_ConOpe.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Concepto Operativo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ConOpe)
      Exit Sub
   End If
   If cmb_DebHab.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Concepto Contable va al Debe o al Haber.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DebHab)
      Exit Sub
   End If
   If cmb_TipCam.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCam)
      Exit Sub
   End If
   
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If l_int_FlgGrb = 1 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
   End If
   
   grd_Listad.Col = 0
   grd_Listad.Text = txt_DesIte.Text
   
   grd_Listad.Col = 1
   grd_Listad.Text = cmb_TipCon.Text
   
   If cmb_TipCon.ItemData(cmb_TipCon.ListIndex) = 1 Then
      grd_Listad.Col = 2
      grd_Listad.Text = ""
      
      grd_Listad.Col = 3
      grd_Listad.Text = Mid(cmb_CtaCtb.Text, 1, l_int_TopNiv)
   Else
      grd_Listad.Col = 2
      grd_Listad.Text = cmb_ConCtb.Text
      
      grd_Listad.Col = 3
      grd_Listad.Text = ""
   End If
   
   grd_Listad.Col = 4
   grd_Listad.Text = cmb_DebHab.Text
   
   grd_Listad.Col = 5
   grd_Listad.Text = cmb_TipCam
   
   grd_Listad.Col = 6
   grd_Listad.Text = cmb_ConOpe.Text
   
   grd_Listad.Col = 7
   grd_Listad.Text = CStr(cmb_TipCon.ItemData(cmb_TipCon.ListIndex))
   
   If cmb_ConCtb.ListIndex > -1 Then
      grd_Listad.Col = 8
      grd_Listad.Text = l_arr_ConCtb(cmb_ConCtb.ListIndex + 1).Genera_Codigo
   End If
   
   grd_Listad.Col = 9
   grd_Listad.Text = l_arr_ConOpe(cmb_ConOpe.ListIndex + 1).Genera_Codigo
   
   grd_Listad.Col = 10
   grd_Listad.Text = CStr(cmb_DebHab.ItemData(cmb_DebHab.ListIndex))
   
   grd_Listad.Col = 11
   grd_Listad.Text = CStr(cmb_TipCam.ItemData(cmb_TipCam.ListIndex))
   
   Call gs_RefrescaGrid(grd_Listad)
   Call cmd_Cancel_Click
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_Activa(False)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_Empres.Caption = moddat_g_str_RazSoc
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_LimpiaItem
   Call fs_Activa(False)
   Call gs_CentraForm(Me)
   
   If moddat_g_int_FlgGrb = 2 Then
      'Cargando Cabecera
      g_str_Parame = "SELECT * FROM CTB_MATCAB WHERE MATCAB_CODMAT = '" & moddat_g_str_Codigo & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   
      g_rst_Princi.MoveFirst
      txt_Codigo.Text = Trim(g_rst_Princi!MATCAB_CODMAT)
      txt_Codigo.Enabled = False
      txt_Descri.Text = Trim(g_rst_Princi!MATCAB_DESCRI)
      Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!MATCAB_TIPMON)
      cmb_TipMat.ListIndex = gf_Busca_Arregl(l_arr_TipMat, Trim(g_rst_Princi!MATCAB_TIPMAT)) - 1
      Call gs_BuscarCombo_Item(cmb_LibCtb, g_rst_Princi!MATCAB_CODLIB)
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      'Cargando Detalle
      g_str_Parame = "SELECT * FROM CTB_MATDET WHERE MATDET_CODMAT = '" & moddat_g_str_Codigo & "' ORDER BY MATDET_NUMITE ASC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = Trim(g_rst_Princi!MATDET_DESCRI & "")
         
         grd_Listad.Col = 1
         grd_Listad.Text = moddat_gf_Consulta_ParDes("062", CStr(g_rst_Princi!MATDET_TIPCON))
      
         grd_Listad.Col = 7
         grd_Listad.Text = CStr(g_rst_Princi!MATDET_TIPCON)
      
         If g_rst_Princi!MATDET_TIPCON = 1 Then
            grd_Listad.Col = 3
            grd_Listad.Text = Trim(g_rst_Princi!MATDET_CTACTB)
         Else
            grd_Listad.Col = 2
            grd_Listad.Text = moddat_gf_Consulta_ConceptoCtb(g_rst_Princi!MATDET_CONCTB)
            
            grd_Listad.Col = 8
            grd_Listad.Text = Trim(g_rst_Princi!MATDET_CONCTB)
         End If
         
         grd_Listad.Col = 4
         grd_Listad.Text = moddat_gf_Consulta_ParDes("255", CStr(g_rst_Princi!MATDET_FLGDHB))
         
         grd_Listad.Col = 10
         grd_Listad.Text = CStr(g_rst_Princi!MATDET_FLGDHB)
   
         grd_Listad.Col = 5
         grd_Listad.Text = moddat_gf_Consulta_ParDes("269", CStr(g_rst_Princi!MATDET_TIPTCA))
   
         grd_Listad.Col = 11
         grd_Listad.Text = CStr(g_rst_Princi!MATDET_TIPTCA)
      
         grd_Listad.Col = 6
         grd_Listad.Text = moddat_gf_Consulta_ParDes("064", Trim(g_rst_Princi!MATDET_CONOPE))
   
         grd_Listad.Col = 9
         grd_Listad.Text = Trim(g_rst_Princi!MATDET_CONOPE)
      
         g_rst_Princi.MoveNext
      Loop
      Call gs_UbiIniGrid(grd_Listad)
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call gs_SetFocus(txt_Descri)
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 2165
   grd_Listad.ColWidth(1) = 1835
   grd_Listad.ColWidth(2) = 3035
   grd_Listad.ColWidth(3) = 1985
   grd_Listad.ColWidth(4) = 965
   grd_Listad.ColWidth(5) = 1755
   grd_Listad.ColWidth(6) = 2805
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   
   l_int_TopNiv = -1
   If moddat_gf_Consulta_ParEmp(l_arr_ParEmp, moddat_g_str_CodEmp, "100", "001") Then
      l_int_TopNiv = l_arr_ParEmp(1).Genera_Cantid
   End If
   
   Call modtac_gs_Carga_CtaCtb(moddat_g_str_CodEmp, cmb_CtaCtb, l_arr_CtaCtb, 0, l_int_TopNiv, -1)
   Call moddat_gs_Carga_LisIte(cmb_TipMat, l_arr_TipMat, 1, "061", 2)
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
   Call moddat_gs_Carga_LibCtb(cmb_LibCtb)
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipCon, 1, "062")
   Call moddat_gs_Carga_LisIte(cmb_ConOpe, l_arr_ConOpe, 1, "064", 2)
   Call moddat_gs_Carga_LisIte_Combo(cmb_DebHab, 1, "255")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipCam, 1, "269")
End Sub

Private Sub fs_Limpia()
   txt_Codigo.Text = ""
   txt_Descri.Text = ""
   cmb_TipMat.ListIndex = -1
   cmb_TipMon.ListIndex = -1
   cmb_LibCtb.ListIndex = -1
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_LimpiaItem()
   txt_DesIte.Text = ""
   cmb_TipCon.ListIndex = -1
   cmb_CtaCtb.ListIndex = -1
   cmb_ConCtb.ListIndex = -1
   cmb_ConOpe.ListIndex = -1
   cmb_DebHab.ListIndex = -1
   cmb_TipCam.ListIndex = -1
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   txt_DesIte.Enabled = p_Habilita
   cmb_TipCon.Enabled = p_Habilita
   cmb_CtaCtb.Enabled = p_Habilita
   cmb_ConCtb.Enabled = p_Habilita
   cmb_ConOpe.Enabled = p_Habilita
   cmb_DebHab.Enabled = p_Habilita
   cmb_TipCam.Enabled = p_Habilita
   cmd_Acepta.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   cmd_Grabar.Enabled = Not p_Habilita
   cmd_Agrega.Enabled = Not p_Habilita
   cmd_Editar.Enabled = Not p_Habilita
   cmd_Borrar.Enabled = Not p_Habilita
   grd_Listad.Enabled = Not p_Habilita
   
   If moddat_g_int_FlgGrb = 2 Then
      txt_Codigo.Enabled = False
   Else
      txt_Codigo.Enabled = Not p_Habilita
   End If
   
   txt_Descri.Enabled = Not p_Habilita
   cmb_TipMat.Enabled = Not p_Habilita
   cmb_TipMon.Enabled = Not p_Habilita
   cmb_LibCtb.Enabled = Not p_Habilita
End Sub


Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub txt_Codigo_GotFocus()
   Call gs_SelecTodo(txt_Codigo)
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descri)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_Descri_GotFocus()
   Call gs_SelecTodo(txt_Descri)
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_TipMat_Click()
   Call gs_SetFocus(cmb_TipMon)
End Sub

Private Sub cmb_TipMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMat_Click
   End If
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(cmb_LibCtb)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmb_LibCtb_Click()
   Call gs_SetFocus(cmd_Agrega)
End Sub

Private Sub cmb_LibCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_LibCtb_Click
   End If
End Sub

Private Sub txt_DesIte_GotFocus()
   Call gs_SelecTodo(txt_DesIte)
End Sub

Private Sub txt_DesIte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipCon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@,.:;_")
   End If
End Sub

Private Sub cmb_TipCon_Click()
   If cmb_TipCon.ListIndex > -1 Then
      If cmb_TipCon.ItemData(cmb_TipCon.ListIndex) = 1 Then
         cmb_CtaCtb.Enabled = True
         cmb_ConCtb.Enabled = False
         cmb_ConCtb.ListIndex = -1
         Call gs_SetFocus(cmb_CtaCtb)
      Else
         cmb_CtaCtb.Enabled = False
         cmb_CtaCtb.ListIndex = -1
         cmb_ConCtb.Enabled = True
         
         Call moddat_gs_Carga_ConceptoCtb(l_arr_ConCtb, cmb_ConCtb, cmb_TipCon.ItemData(cmb_TipCon.ListIndex))
         Call gs_SetFocus(cmb_ConCtb)
      End If
   Else
      Call gs_SetFocus(cmb_ConCtb)
   End If
End Sub

Private Sub cmb_TipCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipCon_Click
   End If
End Sub

Private Sub cmb_CtaCtb_Click()
   If cmb_CtaCtb.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_ConOpe)
      End If
   End If
End Sub

Private Sub cmb_CtaCtb_GotFocus()
   Call SendMessage(cmb_CtaCtb.hWnd, CB_SHOWDROPDOWN, 1, 0&)
   l_int_FlgCmb = True
End Sub

Private Sub cmb_CtaCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_CtaCtb, l_str_CtaCtb)
      l_int_FlgCmb = True
      
      If cmb_CtaCtb.ListIndex > -1 Then
         l_str_CtaCtb = ""
      End If
      
      Call gs_SetFocus(cmb_ConOpe)
   End If
End Sub

Private Sub cmb_CtaCtb_Change()
   l_str_CtaCtb = cmb_CtaCtb.Text
   cmb_CtaCtb.SelLength = Len(l_str_CtaCtb)
End Sub

Private Sub cmb_CtaCtb_LostFocus()
   Call SendMessage(cmb_CtaCtb.hWnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_ConCtb_Click()
   Call gs_SetFocus(cmb_ConOpe)
End Sub

Private Sub cmb_ConCtb_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ConCtb_Click
   End If
End Sub

Private Sub cmb_ConOpe_Click()
   Call gs_SetFocus(cmb_DebHab)
End Sub

Private Sub cmb_ConOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ConOpe_Click
   End If
End Sub

Private Sub cmb_DebHab_Click()
   Call gs_SetFocus(cmb_TipCam)
End Sub

Private Sub cmb_DebHab_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_DebHab_Click
   End If
End Sub

Private Sub cmb_TipCam_Click()
   Call gs_SetFocus(cmd_Acepta)
End Sub

Private Sub cmb_TipCam_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipCam_Click
   End If
End Sub
