VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_PagCom_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15765
   Icon            =   "GesCtb_frm_211.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   15765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7905
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15780
      _Version        =   65536
      _ExtentX        =   27834
      _ExtentY        =   13944
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
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   15620
         _Version        =   65536
         _ExtentX        =   27552
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   495
            Left            =   630
            TabIndex        =   2
            Top             =   90
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Seguimiento de Pagos por Autorizar"
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
            Left            =   30
            Picture         =   "GesCtb_frm_211.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   3
         Top             =   780
         Width           =   15620
         _Version        =   65536
         _ExtentX        =   27552
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
            Picture         =   "GesCtb_frm_211.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_211.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmb_Rechazar 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_211.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Rechazar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3000
            Picture         =   "GesCtb_frm_211.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   15000
            Picture         =   "GesCtb_frm_211.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   2400
            Picture         =   "GesCtb_frm_211.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmb_Aprobar 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_211.frx":17C2
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Aprobar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5475
         Left            =   60
         TabIndex        =   8
         Top             =   2280
         Width           =   15620
         _Version        =   65536
         _ExtentX        =   27552
         _ExtentY        =   9657
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   285
            Left            =   12180
            TabIndex        =   22
            Top             =   60
            Width           =   870
            _Version        =   65536
            _ExtentX        =   1535
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5055
            Left            =   30
            TabIndex        =   9
            Top             =   360
            Width           =   15570
            _ExtentX        =   27464
            _ExtentY        =   8916
            _Version        =   393216
            Rows            =   30
            Cols            =   20
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_NumDoc 
            Height          =   285
            Left            =   5490
            TabIndex        =   17
            Top             =   60
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Documento"
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
         Begin Threed.SSPanel pnl_Fecha 
            Height          =   285
            Left            =   4440
            TabIndex        =   18
            Top             =   60
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha"
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
         Begin Threed.SSPanel pnl_Proveedor 
            Height          =   285
            Left            =   6810
            TabIndex        =   19
            Top             =   60
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5636
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Proveedor"
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   285
            Left            =   60
            TabIndex        =   20
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
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
         Begin Threed.SSPanel pnl_TotPag 
            Height          =   285
            Left            =   13020
            TabIndex        =   21
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2028
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total a Pagar"
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
         Begin Threed.SSPanel pnl_Seleccion 
            Height          =   285
            Left            =   14160
            TabIndex        =   23
            Top             =   60
            Width           =   1080
            _Version        =   65536
            _ExtentX        =   1905
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   " Selección"
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
            Alignment       =   1
            Begin VB.CheckBox chkSeleccionar 
               BackColor       =   &H00004000&
               Caption         =   "Check1"
               Height          =   255
               Left            =   840
               TabIndex        =   24
               Top             =   0
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_TipProducto 
            Height          =   285
            Left            =   1260
            TabIndex        =   25
            Top             =   60
            Width           =   1920
            _Version        =   65536
            _ExtentX        =   3387
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Proceso"
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
         Begin Threed.SSPanel pnl_Usuario 
            Height          =   285
            Left            =   3150
            TabIndex        =   26
            Top             =   60
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Usuario Registro"
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
         Begin Threed.SSPanel pnl_Glosa 
            Height          =   285
            Left            =   9990
            TabIndex        =   27
            Top             =   60
            Width           =   2200
            _Version        =   65536
            _ExtentX        =   3881
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Glosa"
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   60
         TabIndex        =   11
         Top             =   1470
         Width           =   15620
         _Version        =   65536
         _ExtentX        =   27552
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
         Begin VB.CheckBox chk_Estado 
            Caption         =   "Todos los Procesos"
            Height          =   315
            Left            =   6420
            TabIndex        =   16
            Top             =   240
            Width           =   2685
         End
         Begin VB.ComboBox cmb_Proceso 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Proceso:"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   270
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_PagCom_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_bol_FrmIni        As Boolean
Dim l_int_Contar        As Integer

Private Sub chk_Estado_Click()
   If chk_Estado.Value = 1 Then
      cmb_Proceso.ListIndex = -1
      cmb_Proceso.Enabled = False
   Else
      cmb_Proceso.Enabled = True
   End If
End Sub

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 9) = ""
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 9) = "X"
         Next r_Fila
      End If
   Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub cmb_Aprobar_Click()
Dim r_bol_Estado    As Boolean
Dim r_int_Fila      As Integer
Dim r_str_CodGrb    As String

   moddat_g_str_Codigo = ""
   r_str_CodGrb = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   r_bol_Estado = False
   For r_int_Fila = 0 To grd_Listad.Rows - 1
       If Trim(grd_Listad.TextMatrix(r_int_Fila, 9)) = "X" Then
          r_bol_Estado = True
          Exit For
       End If
   Next
   
   If r_bol_Estado = False Then
      MsgBox "No hay ninguna fila seleccionada.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   Call gs_RefrescaGrid(grd_Listad)
   If MsgBox("¿Seguro que desea aprobar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   For r_int_Fila = 0 To grd_Listad.Rows - 1
       If Trim(grd_Listad.TextMatrix(r_int_Fila, 9)) = "X" Then
          g_str_Parame = ""
          g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT_ESTADO ( "
          g_str_Parame = g_str_Parame & "'" & CLng(grd_Listad.TextMatrix(r_int_Fila, 10)) & "', " 'COMAUT_CODAUT
          g_str_Parame = g_str_Parame & " " & 2 & ", " 'APROBADO
          g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
          g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
          g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
          g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
         
          If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
          If g_rst_Genera!RESUL = 1 Then
             r_str_CodGrb = r_str_CodGrb & " - " & CStr(grd_Listad.TextMatrix(r_int_Fila, 0)) 'COMAUT_CODOPE
          End If
       End If
   Next
   
   MsgBox "Registros aprobados correctamente." & vbCrLf & "Codigos :" & r_str_CodGrb, vbInformation, modgen_g_str_NomPlt
   
   Screen.MousePointer = 0
   Call fs_Buscar
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmb_Rechazar_Click()
Dim r_bol_Estado  As Boolean
Dim r_int_Fila    As Integer
Dim r_str_CodGrb  As String

   moddat_g_str_Codigo = ""
   r_str_CodGrb = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
         
   r_bol_Estado = False
   For r_int_Fila = 0 To grd_Listad.Rows - 1
       If Trim(grd_Listad.TextMatrix(r_int_Fila, 9)) = "X" Then
          r_bol_Estado = True
          Exit For
       End If
   Next
   
   If r_bol_Estado = False Then
      MsgBox "No hay ninguna fila seleccionada.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   Call gs_RefrescaGrid(grd_Listad)
   If MsgBox("¿Seguro que desea rechazar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   For r_int_Fila = 0 To grd_Listad.Rows - 1
       If Trim(grd_Listad.TextMatrix(r_int_Fila, 9)) = "X" Then
          g_str_Parame = ""
          g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT_ESTADO ( "
          g_str_Parame = g_str_Parame & " " & CLng(grd_Listad.TextMatrix(r_int_Fila, 10)) & ", " 'COMAUT_CODAUT
          g_str_Parame = g_str_Parame & " " & 3 & ", " 'RECHAZAR
          g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
          g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
          g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
          g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
          If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
          If g_rst_Genera!RESUL = 1 Then
             r_str_CodGrb = r_str_CodGrb & " - " & CStr(grd_Listad.TextMatrix(r_int_Fila, 0)) 'COMAUT_CODOPE
          End If
       End If
   Next
   
   MsgBox "Registros rechazados correctamente." & vbCrLf & "Codigos :" & r_str_CodGrb, vbInformation, modgen_g_str_NomPlt
   
   Screen.MousePointer = 0
   Call fs_Buscar
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Buscar_Click()
   If l_bol_FrmIni = True Then
      If chk_Estado.Value = 0 Then
         If cmb_Proceso.ListIndex = -1 Then
            MsgBox "Seleccione un tipo de proceso.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Proceso)
            Exit Sub
         End If
      End If
   End If
   
   Call fs_Buscar
   chk_Estado.Enabled = False
   cmb_Proceso.Enabled = False
End Sub

Private Sub cmd_Consul_Click()
Dim r_str_CodAux   As String
Dim r_str_FlgAux   As Integer

   r_str_CodAux = ""
   r_str_FlgAux = 0
   moddat_g_str_NumOpe = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   Call gs_RefrescaGrid(grd_Listad)
   
   Select Case Left(grd_Listad.TextMatrix(grd_Listad.Row, 0), 2)
          Case "01" 'CUENTAS X PAGAR
               moddat_g_str_NumOpe = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
               frm_Ctb_PagCom_04.Show 1
          Case "12" 'CUENTAS X PAGAR GESCTB
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  frm_Ctb_CtaPag_02.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "07" 'GESTION PERSONAL
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  moddat_g_int_TipRec = 1 'GESTION DE PAGOS
                  frm_Ctb_GesPer_02.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "08" 'CARGA DEL ARCHIVO RECAUDO
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  frm_Ctb_CarArc_02.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "06" 'REGISTRO DE COMPRAS
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_int_FlgGrb = 0
                  moddat_g_str_TipDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 11))
                  moddat_g_str_NumDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 12))
                  moddat_g_int_InsAct = 0 'tipo registro compra
                  frm_Ctb_RegCom_04.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "05" 'ENTREGAS A RENDIR
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_str_CodIte = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_str_CodMod = grd_Listad.TextMatrix(grd_Listad.Row, 13)
                  moddat_g_int_FlgGrb = 0
                  If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 19) & "") = "1" Then
                     frm_Ctb_EntRen_02.Show 1 'form principal
                  ElseIf Trim(grd_Listad.TextMatrix(grd_Listad.Row, 19) & "") = "2" Then
                     frm_Ctb_EntRen_04.Show 1 'reembolso
                  End If
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case Else
               Exit Sub
   End Select
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   chk_Estado.Enabled = True
   chk_Estado.Value = 0
   Call chk_Estado_Click
   Call gs_SetFocus(chk_Estado)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   l_bol_FrmIni = False
   
   Call fs_Inicia
   chk_Estado.Value = 1
   Call cmd_Buscar_Click
   
   l_bol_FrmIni = True
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1200 'CODIGO
   grd_Listad.ColWidth(1) = 1900 'TIPO PROCESO
   grd_Listad.ColWidth(2) = 1280 'USUARIO REGISTRO
   grd_Listad.ColWidth(3) = 1050 'FECHA
   grd_Listad.ColWidth(4) = 1340 'NRO-DOCUMENTO
   grd_Listad.ColWidth(5) = 3170 'PROVEEDOR
   grd_Listad.ColWidth(6) = 2190 'GLOSA
   grd_Listad.ColWidth(7) = 850 'MONEDA
   grd_Listad.ColWidth(8) = 1110 'TOTAL
   grd_Listad.ColWidth(9) = 1090 'SELECCIONAR
   grd_Listad.ColWidth(10) = 0 'COMAUT_CODAUT
   grd_Listad.ColWidth(11) = 0 'COMAUT_TIPDOC
   grd_Listad.ColWidth(12) = 0 'COMAUT_NUMDOC
   grd_Listad.ColWidth(13) = 0 'COMAUT_CODMON
   grd_Listad.ColWidth(14) = 0 'COMAUT_CODBNC
   grd_Listad.ColWidth(15) = 0 'COMAUT_CTACTB
   grd_Listad.ColWidth(16) = 0 'COMAUT_DATCTB
   grd_Listad.ColWidth(17) = 0 'COMAUT_FECHA - ORDEN
   grd_Listad.ColWidth(18) = 0 'COMAUT_IMPORTE -ORDEN
   grd_Listad.ColWidth(19) = 0 'COMAUT_TIPOPE
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignLeftCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_Proceso, 1, "136")
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Buscar()
Dim r_str_FecIni  As String
Dim r_str_FecFin  As String
Dim r_str_Cadena  As String

   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT COMAUT_CODAUT, COMAUT_CODOPE, COMAUT_FECOPE, COMAUT_TIPDOC, COMAUT_NUMDOC, "
   g_str_Parame = g_str_Parame & "        COMAUT_CODMON, COMAUT_IMPPAG, COMAUT_CODBNC, COMAUT_CTACRR, COMAUT_CTACTB, "
   g_str_Parame = g_str_Parame & "        COMAUT_DATCTB , COMAUT_CODEST, TRIM(C.PARDES_DESCRI) AS MONEDA, COMAUT_USUINI, "
   g_str_Parame = g_str_Parame & "        TRIM(D.PARDES_DESCRI) AS TIPOPROCESO, TRIM(A.COMAUT_DESCRP) AS GLOSA, COMAUT_TIPOPE,  "
   g_str_Parame = g_str_Parame & "        DECODE(B.MaePrv_RazSoc,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE) "
   g_str_Parame = g_str_Parame & "               ,B.MaePrv_RazSoc) AS MaePrv_RazSoc "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_COMAUT A  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.COMAUT_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(A.COMAUT_NUMDOC) "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = A.COMAUT_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(A.COMAUT_NUMDOC) "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.COMAUT_CODMON "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 136 AND TO_NUMBER(D.PARDES_CODITE) = TO_NUMBER(SUBSTR(LPAD(COMAUT_CODOPE,10,0),1,2)) AND D.PARDES_CODITE <> 0 "
   g_str_Parame = g_str_Parame & "  WHERE COMAUT_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "    AND COMAUT_CODEST = 1 "
   If chk_Estado.Value = 0 And cmb_Proceso.ListIndex <> -1 Then
      g_str_Parame = g_str_Parame & "    AND TO_NUMBER(SUBSTR(LPAD(COMAUT_CODOPE,10,0),1,2)) = " & CStr(cmb_Proceso.ItemData(cmb_Proceso.ListIndex))
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY COMAUT_FECOPE ASC, A.COMAUT_CODOPE ASC  "
   '---------------------------------------
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If

   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst

   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1

      grd_Listad.Col = 0
      grd_Listad.Text = Format(Trim(g_rst_Princi!COMAUT_CODOPE), "0000000000")

      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!TIPOPROCESO & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!COMAUT_USUINI & "")
      
      grd_Listad.Col = 3
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!COMAUT_FECOPE)

      grd_Listad.Col = 4
      grd_Listad.Text = g_rst_Princi!COMAUT_TIPDOC & "-" & Trim(g_rst_Princi!COMAUT_NUMDOC & "")
      
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!MaePrv_RazSoc & "")
            
      grd_Listad.Col = 6
      grd_Listad.Text = CStr(g_rst_Princi!GLOSA & "")
      
      
      grd_Listad.Col = 7
      grd_Listad.Text = CStr(g_rst_Princi!Moneda & "")
            
      grd_Listad.Col = 8 'TOTAL A PAGAR
      grd_Listad.Text = Format(g_rst_Princi!COMAUT_IMPPAG, "###,###,###,##0.00")
                                 
      'grd_Listad.Col = 8
      'grd_Listad.Text = COLUMNA SELECCIONAR
                                 
      grd_Listad.Col = 10
      grd_Listad.Text = g_rst_Princi!COMAUT_CODAUT
      
      grd_Listad.Col = 11
      grd_Listad.Text = g_rst_Princi!COMAUT_TIPDOC
      
      grd_Listad.Col = 12
      grd_Listad.Text = g_rst_Princi!COMAUT_NUMDOC
      
      grd_Listad.Col = 13
      grd_Listad.Text = g_rst_Princi!COMAUT_CODMON
      
      grd_Listad.Col = 14
      grd_Listad.Text = Trim(CStr(g_rst_Princi!COMAUT_CODBNC & ""))
      
      grd_Listad.Col = 15
      grd_Listad.Text = g_rst_Princi!COMAUT_CTACTB
      
      grd_Listad.Col = 16
      grd_Listad.Text = g_rst_Princi!COMAUT_DATCTB
      '---------------------
      grd_Listad.Col = 17
      grd_Listad.Text = g_rst_Princi!COMAUT_FECOPE
      
      grd_Listad.Col = 18
      grd_Listad.Text = g_rst_Princi!COMAUT_IMPPAG
      
      grd_Listad.Col = 19
      grd_Listad.Text = g_rst_Princi!COMAUT_TIPOPE
      
      g_rst_Princi.MoveNext
   Loop

   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 9
      If grd_Listad.Text = "X" Then
          grd_Listad.Text = ""
      Else
           grd_Listad.Text = "X"
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE PAGOS POR AUTORIZAR"
      .Range(.Cells(2, 2), .Cells(2, 10)).Merge
      .Range(.Cells(2, 2), .Cells(2, 10)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 10)).HorizontalAlignment = xlHAlignCenter

      .Cells(3, 2) = "CÓDIGO"
      .Cells(3, 3) = "TIPO PROCESO"
      .Cells(3, 4) = "USUARIO REGISTRO"
      .Cells(3, 5) = "FECHA"
      .Cells(3, 6) = "NRO DOCUMENTO"
      .Cells(3, 7) = "PROVEEDOR"
      .Cells(3, 8) = "GLOSA"
      .Cells(3, 9) = "MONEDA"
      .Cells(3, 10) = "TOTAL A PAGAR"
         
      .Range(.Cells(3, 2), .Cells(3, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 10)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13 'codigo
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 22 'tipo proceso
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 22 'USUARIO REGISTRO
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 12 'FECHA
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 18 'NRO DOCUMENTO
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 44 'PROVEEDOR
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 30 'GLOSA
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 22 'MONEDA
      .Columns("I").HorizontalAlignment = xlHAlignLeft
      .Columns("J").ColumnWidth = 17 'TOTAL A PAGAR
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
            
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Size = 11
      
      r_int_NumFil = 4
      For l_int_Contar = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil, 2) = "'" & grd_Listad.TextMatrix(l_int_Contar, 0)
         .Cells(r_int_NumFil, 3) = grd_Listad.TextMatrix(l_int_Contar, 1)
         .Cells(r_int_NumFil, 4) = grd_Listad.TextMatrix(l_int_Contar, 2)
         .Cells(r_int_NumFil, 5) = "'" & grd_Listad.TextMatrix(l_int_Contar, 3)
         .Cells(r_int_NumFil, 6) = grd_Listad.TextMatrix(l_int_Contar, 4)
         .Cells(r_int_NumFil, 7) = "'" & grd_Listad.TextMatrix(l_int_Contar, 5)
         .Cells(r_int_NumFil, 8) = "'" & grd_Listad.TextMatrix(l_int_Contar, 6)
         .Cells(r_int_NumFil, 9) = "'" & grd_Listad.TextMatrix(l_int_Contar, 7)
         .Cells(r_int_NumFil, 10) = grd_Listad.TextMatrix(l_int_Contar, 8)
         
         r_int_NumFil = r_int_NumFil + 1
      Next
      .Range(.Cells(3, 3), .Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
      
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub pnl_Codigo_Click()
   If pnl_Codigo.Tag = "" Then
      Call gs_SorteaGrid(grd_Listad, 0, "N")
      pnl_Codigo.Tag = 1
   Else
      Call gs_SorteaGrid(grd_Listad, 0, "N-")
      pnl_Codigo.Tag = ""
   End If
End Sub

Private Sub pnl_Glosa_Click()
   If pnl_Glosa.Tag = "" Then
      Call gs_SorteaGrid(grd_Listad, 6, "C")
      pnl_Glosa.Tag = 1
   Else
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
      pnl_Glosa.Tag = ""
   End If
End Sub

Private Sub pnl_Moneda_Click()
   If pnl_Moneda.Tag = "" Then
      Call gs_SorteaGrid(grd_Listad, 7, "C")
      pnl_Moneda.Tag = 1
   Else
      Call gs_SorteaGrid(grd_Listad, 7, "C-")
      pnl_Moneda.Tag = ""
   End If
End Sub

Private Sub pnl_NumDoc_Click()
   If pnl_NumDoc.Tag = "" Then
      Call gs_SorteaGrid(grd_Listad, 4, "N")
      pnl_NumDoc.Tag = 1
   Else
      Call gs_SorteaGrid(grd_Listad, 4, "N-")
      pnl_NumDoc.Tag = ""
   End If
End Sub

Private Sub pnl_Proveedor_Click()
   If pnl_Proveedor.Tag = "" Then
      Call gs_SorteaGrid(grd_Listad, 5, "C")
      pnl_Proveedor.Tag = 1
   Else
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
      pnl_Proveedor.Tag = ""
   End If
End Sub

Private Sub pnl_Seleccion_Click()
   If pnl_Seleccion.Tag = "" Then
      Call gs_SorteaGrid(grd_Listad, 9, "C")
      pnl_Seleccion.Tag = 1
   Else
      Call gs_SorteaGrid(grd_Listad, 9, "C-")
      pnl_Seleccion.Tag = ""
   End If
End Sub

Private Sub pnl_TipProducto_Click()
   If pnl_TipProducto.Tag = "" Then
      Call gs_SorteaGrid(grd_Listad, 1, "C")
      pnl_TipProducto.Tag = 1
   Else
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
      pnl_TipProducto.Tag = ""
   End If
End Sub

Private Sub pnl_Usuario_Click()
   If pnl_Usuario.Tag = "" Then
      Call gs_SorteaGrid(grd_Listad, 2, "C")
      pnl_Usuario.Tag = 1
   Else
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
      pnl_Usuario.Tag = ""
   End If
End Sub

Private Sub pnl_Fecha_Click()
   If pnl_Fecha.Tag = "" Then
      Call gs_SorteaGrid(grd_Listad, 17, "N")
      pnl_Fecha.Tag = 1
   Else
      Call gs_SorteaGrid(grd_Listad, 17, "N-")
      pnl_Fecha.Tag = ""
   End If
End Sub

Private Sub pnl_TotPag_Click()
   If pnl_TotPag.Tag = "" Then
      Call gs_SorteaGrid(grd_Listad, 18, "N")
      pnl_TotPag.Tag = 1
   Else
      Call gs_SorteaGrid(grd_Listad, 18, "N-")
      pnl_TotPag.Tag = ""
   End If
End Sub


