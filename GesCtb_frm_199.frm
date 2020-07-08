VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_InvDpf_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13815
   Icon            =   "GesCtb_frm_199.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8415
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   14145
      _Version        =   65536
      _ExtentX        =   24950
      _ExtentY        =   14843
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
         TabIndex        =   15
         Top             =   60
         Width           =   13720
         _Version        =   65536
         _ExtentX        =   24201
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
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de Cuentas - Depósito Plazo Fijo"
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
            Picture         =   "GesCtb_frm_199.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   4005
         Left            =   60
         TabIndex        =   17
         Top             =   2490
         Width           =   13720
         _Version        =   65536
         _ExtentX        =   24201
         _ExtentY        =   7064
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
         Begin Threed.SSPanel pnl_Tit_Descri 
            Height          =   285
            Left            =   2775
            TabIndex        =   18
            Top             =   90
            Width           =   2730
            _Version        =   65536
            _ExtentX        =   4815
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Origen de Fondos"
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
            TabIndex        =   19
            Top             =   90
            Width           =   2730
            _Version        =   65536
            _ExtentX        =   4815
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Transacción"
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
            Left            =   5490
            TabIndex        =   25
            Top             =   90
            Width           =   2940
            _Version        =   65536
            _ExtentX        =   5186
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Rubro"
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
            Left            =   9290
            TabIndex        =   26
            Top             =   90
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2417
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Debe"
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
            Left            =   10635
            TabIndex        =   27
            Top             =   90
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2417
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Haber -1"
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
            Left            =   11985
            TabIndex        =   28
            Top             =   90
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2417
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Haber -2"
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   3555
            Left            =   60
            TabIndex        =   29
            Top             =   390
            Width           =   13640
            _ExtentX        =   24051
            _ExtentY        =   6271
            _Version        =   393216
            Rows            =   80
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   8400
            TabIndex        =   36
            Top             =   90
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
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
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   20
         Top             =   780
         Width           =   13720
         _Version        =   65536
         _ExtentX        =   24201
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
            Left            =   60
            Picture         =   "GesCtb_frm_199.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_199.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_199.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_199.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Editar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1260
            Picture         =   "GesCtb_frm_199.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13110
            Picture         =   "GesCtb_frm_199.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   3060
            Picture         =   "GesCtb_frm_199.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   3660
            Picture         =   "GesCtb_frm_199.frx":1ACC
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Cancelar "
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   1755
         Left            =   60
         TabIndex        =   21
         Top             =   6540
         Width           =   13720
         _Version        =   65536
         _ExtentX        =   24201
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
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   390
            Width           =   1995
         End
         Begin VB.ComboBox cmb_Rubro 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   60
            Width           =   5000
         End
         Begin VB.TextBox txt_Haber_02 
            Height          =   315
            Left            =   1620
            MaxLength       =   12
            TabIndex        =   13
            Top             =   1380
            Width           =   1965
         End
         Begin VB.TextBox txt_Haber_01 
            Height          =   315
            Left            =   1620
            MaxLength       =   12
            TabIndex        =   12
            Top             =   1050
            Width           =   1965
         End
         Begin VB.TextBox txt_Debe_01 
            Height          =   315
            Left            =   1620
            MaxLength       =   12
            TabIndex        =   11
            Top             =   720
            Width           =   1965
         End
         Begin VB.Label lbl_Etique 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   35
            Top             =   450
            Width           =   630
         End
         Begin VB.Label lbl_Etique 
            AutoSize        =   -1  'True
            Caption         =   "Rubro:"
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   33
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lbl_Etique 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Haber-2:"
            Height          =   195
            Index           =   4
            Left            =   150
            TabIndex        =   24
            Top             =   1440
            Width           =   1170
         End
         Begin VB.Label lbl_Etique 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Haber-1:"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   23
            Top             =   1110
            Width           =   1170
         End
         Begin VB.Label lbl_Etique 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Debe:"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   22
            Top             =   780
            Width           =   990
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   975
         Left            =   60
         TabIndex        =   30
         Top             =   1470
         Width           =   13720
         _Version        =   65536
         _ExtentX        =   24201
         _ExtentY        =   1720
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
         Begin VB.ComboBox cmb_Trans 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   5000
         End
         Begin VB.ComboBox cmb_Fondos 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   510
            Width           =   5000
         End
         Begin VB.Label lbl_Etique 
            AutoSize        =   -1  'True
            Caption         =   "Origen de fondos:"
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   32
            Top             =   540
            Width           =   1260
         End
         Begin VB.Label lbl_Etique 
            AutoSize        =   -1  'True
            Caption         =   "Transacción:"
            Height          =   195
            Index           =   6
            Left            =   150
            TabIndex        =   31
            Top             =   240
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_InvDpf_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Screen.MousePointer = 11
   Call fs_Activa2(False)
   Call fs_Activa(True)
   cmb_Trans.Enabled = False
   cmb_Fondos.Enabled = False
   cmd_Buscar.Enabled = False
   Call fs_Limpia
   Call gs_SetFocus(cmb_Rubro)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CTB_CTADPF ( "
   g_str_Parame = g_str_Parame & CLng(cmb_Trans.ItemData(cmb_Trans.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 7)) & ", "
   g_str_Parame = g_str_Parame & CLng(cmb_Fondos.ItemData(cmb_Fondos.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 8)) & ", "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Debe_01.Text) & "', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Haber_01.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Haber_02.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & "3) " 'as_insupd
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar el borrado de datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If (g_rst_Genera!RESUL = 1) Then
       MsgBox "Los datos se borraron correctamente.", vbInformation, modgen_g_str_NomPlt
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Activa2(True)
   Call fs_Activa(False)
   Call fs_Limpia
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_int_FlgGrb = 2
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   Call fs_Activa2(False)
   Call fs_Activa(True)
   cmb_Trans.Enabled = False
   cmb_Fondos.Enabled = False
   cmd_Buscar.Enabled = False
   Call fs_Limpia
   
   Call gs_BuscarCombo_Item(cmb_Rubro, Trim(grd_Listad.TextMatrix(grd_Listad.Row, 7)))
   Call gs_BuscarCombo_Item(cmb_Moneda, Trim(grd_Listad.TextMatrix(grd_Listad.Row, 8)))
   
   txt_Debe_01.Text = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 4))
   txt_Haber_01.Text = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 5))
   txt_Haber_02.Text = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 6))
   
   Call gs_SetFocus(cmb_Rubro)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()
Dim r_int_Contad   As Integer
Dim r_bol_Estado   As Boolean

  If cmb_Trans.ListIndex = -1 Then
      MsgBox "Debe de seleccionar el tipo de transacción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Trans)
      Exit Sub
   End If
   If cmb_Fondos.ListIndex = -1 Then
      MsgBox "Debe de seleccionar el tipo de fondo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Fondos)
      Exit Sub
   End If
   If cmb_Rubro.ListIndex = -1 Then
      MsgBox "Debe de seleccionar el tipo de rubro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Rubro)
      Exit Sub
   End If
   If cmb_Moneda.ListIndex = -1 Then
      MsgBox "Debe de seleccionar el tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Moneda)
      Exit Sub
   End If
   
   If Len(Trim(txt_Debe_01.Text)) = 0 Then
      MsgBox "Debe de ingresar la cuentra contable en el debe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Debe_01)
      Exit Sub
   End If
   If Len(Trim(txt_Haber_01.Text)) = 0 Then
      MsgBox "Debe de ingresar la cuentra contable en el haber-1.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Haber_01)
      Exit Sub
   End If
   'If Len(Trim(txt_Haber_02.Text)) = 0 Then
   '   MsgBox "Debe de ingresar la cuentra contable en el haber-2.", vbExclamation, modgen_g_str_NomPlt
   '   Call gs_SetFocus(txt_Haber_02)
   '   Exit Sub
   'End If
   
   r_bol_Estado = True
   For r_int_Contad = 0 To grd_Listad.Rows - 1
       If moddat_g_int_FlgGrb = 1 Then 'INSERT
          If CInt(cmb_Rubro.ItemData(cmb_Rubro.ListIndex)) = CInt(grd_Listad.TextMatrix(r_int_Contad, 7)) And _
             CInt(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)) = CInt(grd_Listad.TextMatrix(r_int_Contad, 8)) Then
             r_bol_Estado = False
          End If
       Else 'EDITAR
          If grd_Listad.Row <> r_int_Contad Then
             If CInt(cmb_Rubro.ItemData(cmb_Rubro.ListIndex)) = CInt(grd_Listad.TextMatrix(r_int_Contad, 7)) And _
                CInt(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)) = CInt(grd_Listad.TextMatrix(r_int_Contad, 8)) Then
                r_bol_Estado = False
             End If
          End If
       End If
   Next
   If r_bol_Estado = False Then
      MsgBox "El Rubro ya ha sido agregado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Rubro)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucción SQL
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CTB_CTADPF ( "
   g_str_Parame = g_str_Parame & CLng(cmb_Trans.ItemData(cmb_Trans.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & CLng(cmb_Rubro.ItemData(cmb_Rubro.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & CLng(cmb_Fondos.ItemData(cmb_Fondos.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & CLng(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Debe_01.Text) & "', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Haber_01.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Haber_02.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & moddat_g_int_FlgGrb & ") " 'as_insupd

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar el grabado de datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If (g_rst_Genera!RESUL = 1) Then
       MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
   ElseIf (g_rst_Genera!RESUL = 2) Then
       MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
   End If
   
   Call fs_Buscar
   Call fs_Activa2(True)
   Call fs_Activa(False)
   Call fs_Limpia
End Sub

Private Sub cmd_Limpia_Click()
   cmb_Trans.ListIndex = -1
   cmb_Rubro.ListIndex = -1
   cmb_Fondos.ListIndex = -1
   cmb_Moneda.ListIndex = -1
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Activa2(True)
   Call fs_Activa(True)
   Call fs_Limpia
   Call gs_SetFocus(cmb_Trans)
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Trans)
   Screen.MousePointer = 0
End Sub
 
 Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 2680
   grd_Listad.ColWidth(1) = 2720
   grd_Listad.ColWidth(2) = 2920
   grd_Listad.ColWidth(3) = 870
   grd_Listad.ColWidth(4) = 1360
   grd_Listad.ColWidth(5) = 1330
   grd_Listad.ColWidth(6) = 1360
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Trans, 1, "122")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Rubro, 1, "128")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Fondos, 1, "122")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_Trans.ListIndex = -1 Then
      MsgBox "Tiene que seleccione el tipo de transacción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Trans)
      Exit Sub
   End If

   If cmb_Fondos.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar el origen de fondo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Fondos)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Activa2(True)
   Call fs_Activa(False)
   Call fs_Buscar
   Call fs_Limpia
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_Trans.Enabled = p_Activa
   'cmb_Rubro.Enabled = p_Activa
   cmb_Fondos.Enabled = p_Activa
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
   
   cmb_Rubro.Enabled = Not p_Activa
   cmb_Moneda.Enabled = Not p_Activa
   txt_Debe_01.Enabled = Not p_Activa
   txt_Haber_01.Enabled = Not p_Activa
   txt_Haber_02.Enabled = Not p_Activa
End Sub

Private Sub fs_Limpia()
   cmb_Rubro.ListIndex = -1
   cmb_Moneda.ListIndex = -1
   txt_Debe_01.Text = ""
   txt_Haber_01.Text = ""
   txt_Haber_02.Text = ""
End Sub

Public Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT TRIM(B.PARDES_DESCRI) AS NOM_TRANS, TRIM(C.PARDES_DESCRI) AS NOM_FONDO,  "
   g_str_Parame = g_str_Parame & "        TRIM(D.PARDES_DESCRI) AS NOM_RUBRO, A.CTADPF_TIPDPF, A.CTADPF_CTADEB_01,  "
   g_str_Parame = g_str_Parame & "        A.CTADPF_CTADEB_02, A.CTADPF_CTAHAB_01, A.CTADPF_CTAHAB_02, A.CTADPF_CODMON,  "
   g_str_Parame = g_str_Parame & "        TRIM(E.PARDES_DESCRI) AS NOM_MONEDA  "
   g_str_Parame = g_str_Parame & "   FROM CTB_CTADPF A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES B ON A.CTADPF_CODENT_DES = B.PARDES_CODITE AND B.PARDES_CODGRP = 122  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON A.CTADPF_CODENT_ORI = C.PARDES_CODITE AND C.PARDES_CODGRP = 122  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON A.CTADPF_TIPDPF = D.PARDES_CODITE AND D.PARDES_CODGRP = 128  "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 204 AND A.CTADPF_CODMON = E.PARDES_CODITE " 'moneda
   
   g_str_Parame = g_str_Parame & "  WHERE A.CTADPF_CODENT_DES =  " & CLng(cmb_Trans.ItemData(cmb_Trans.ListIndex))
   'g_str_Parame = g_str_Parame & "    AND A.CTADPF_TIPDPF =  " & CLng(cmb_Rubro.ItemData(cmb_Rubro.ListIndex))
   g_str_Parame = g_str_Parame & "    AND A.CTADPF_CODENT_ORI =  " & CLng(cmb_Fondos.ItemData(cmb_Fondos.ListIndex))
   g_str_Parame = g_str_Parame & "  ORDER BY A.CTADPF_CODMON,A.CTADPF_TIPDPF ASC  "
        
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
      grd_Listad.Text = Trim(g_rst_Princi!NOM_TRANS & "")
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!NOM_FONDO & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(g_rst_Princi!NOM_RUBRO & "")
                     
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!NOM_MONEDA & "")
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!CTADPF_CTADEB_01 & "")
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!CTADPF_CTAHAB_01 & "")
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!CTADPF_CTAHAB_02 & "")
      
      grd_Listad.Col = 7
      grd_Listad.Text = Trim(g_rst_Princi!CTADPF_TIPDPF & "")
      
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(g_rst_Princi!CTADPF_CODMON & "")
      
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

Private Sub cmb_Trans_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Fondos)
   End If
End Sub

Private Sub cmb_Rubro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Moneda)
   End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Debe_01)
   End If
End Sub

Private Sub cmb_Fondos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub txt_Debe_01_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Haber_01)
   End If
End Sub

Private Sub txt_Haber_01_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Haber_02)
   End If
End Sub

Private Sub txt_Haber_02_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub


