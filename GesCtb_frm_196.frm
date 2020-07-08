VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Ctb_InvDpf_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16080
   Icon            =   "GesCtb_frm_196.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9015
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   16305
      _Version        =   65536
      _ExtentX        =   28760
      _ExtentY        =   15901
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
         TabIndex        =   18
         Top             =   60
         Width           =   15975
         _Version        =   65536
         _ExtentX        =   28178
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
            Left            =   570
            TabIndex        =   19
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Inversiones de Depósito Plazo Fijo"
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
            Picture         =   "GesCtb_frm_196.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   20
         Top             =   780
         Width           =   15975
         _Version        =   65536
         _ExtentX        =   28178
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
         Begin VB.CommandButton cmd_Reversa 
            Height          =   585
            Left            =   3660
            Picture         =   "GesCtb_frm_196.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Reversa"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Generar 
            Height          =   585
            Left            =   5490
            Picture         =   "GesCtb_frm_196.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Generar Asientos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_196.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_196.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Detalle 
            Height          =   585
            Left            =   4860
            Picture         =   "GesCtb_frm_196.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Detalle"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Gestion 
            Height          =   585
            Left            =   4260
            Picture         =   "GesCtb_frm_196.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Gestionar Inversión"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_196.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_196.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_196.frx":1C9E
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   15360
            Picture         =   "GesCtb_frm_196.frx":1FA8
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   3060
            Picture         =   "GesCtb_frm_196.frx":23EA
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   6060
            Picture         =   "GesCtb_frm_196.frx":26F4
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   630
         Left            =   60
         TabIndex        =   21
         Top             =   8220
         Width           =   15975
         _Version        =   65536
         _ExtentX        =   28178
         _ExtentY        =   1111
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
         Begin VB.TextBox txt_Buscar 
            Height          =   315
            Left            =   5160
            MaxLength       =   100
            TabIndex        =   16
            Top             =   180
            Width           =   4155
         End
         Begin VB.ComboBox cmb_Buscar 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   180
            Width           =   2595
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Columna a Buscar:"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Por:"
            Height          =   195
            Left            =   4290
            TabIndex        =   22
            Top             =   240
            Width           =   825
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   5835
         Left            =   60
         TabIndex        =   24
         Top             =   2340
         Width           =   15975
         _Version        =   65536
         _ExtentX        =   28178
         _ExtentY        =   10292
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
            Height          =   5415
            Left            =   30
            TabIndex        =   25
            Top             =   360
            Width           =   15890
            _ExtentX        =   28019
            _ExtentY        =   9551
            _Version        =   393216
            Rows            =   24
            Cols            =   19
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_DebMN 
            Height          =   285
            Left            =   5370
            TabIndex        =   26
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Plazo Dias"
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
         Begin Threed.SSPanel pnl_HabME 
            Height          =   285
            Left            =   6240
            TabIndex        =   27
            Top             =   60
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1402
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tasa %"
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   7020
            TabIndex        =   29
            Top             =   60
            Width           =   800
            _Version        =   65536
            _ExtentX        =   1411
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   7800
            TabIndex        =   30
            Top             =   60
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Capital"
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
            Left            =   1080
            TabIndex        =   31
            Top             =   60
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Institución"
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
            Left            =   3570
            TabIndex        =   32
            Top             =   60
            Width           =   1810
            _Version        =   65536
            _ExtentX        =   3193
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Operación de Ref."
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1887
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Cuenta"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   10320
            TabIndex        =   34
            Top             =   60
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1976
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vencimiento"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   11430
            TabIndex        =   35
            Top             =   60
            Width           =   1100
            _Version        =   65536
            _ExtentX        =   1940
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Rendimeinto"
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   12510
            TabIndex        =   36
            Top             =   60
            Width           =   1150
            _Version        =   65536
            _ExtentX        =   2028
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Devengado"
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   9195
            TabIndex        =   28
            Top             =   60
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Apertura"
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
         Begin Threed.SSPanel SSPanel17 
            Height          =   285
            Left            =   13650
            TabIndex        =   43
            Top             =   60
            Width           =   850
            _Version        =   65536
            _ExtentX        =   1499
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Dias Trans"
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
         Begin Threed.SSPanel SSPanel18 
            Height          =   285
            Left            =   14490
            TabIndex        =   44
            Top             =   60
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Estado"
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
      Begin Threed.SSPanel SSPanel12 
         Height          =   825
         Left            =   60
         TabIndex        =   37
         Top             =   1470
         Width           =   15975
         _Version        =   65536
         _ExtentX        =   28178
         _ExtentY        =   1455
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
            Caption         =   "Todos los Estados"
            Height          =   195
            Left            =   10200
            TabIndex        =   2
            Top             =   510
            Width           =   1875
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   3465
         End
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   3465
         End
         Begin Threed.SSPanel pnl_Period 
            Height          =   315
            Left            =   6780
            TabIndex        =   38
            Top             =   90
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Período Vigente:"
            Height          =   195
            Index           =   2
            Left            =   5310
            TabIndex        =   42
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   41
            Top             =   120
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal:"
            Height          =   195
            Left            =   180
            TabIndex        =   40
            Top             =   450
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Apertura:"
            Height          =   195
            Left            =   5310
            TabIndex        =   39
            Top             =   450
            Width           =   1140
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_InvDpf_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim l_arr_Empres()  As moddat_tpo_Genera
Dim l_arr_Sucurs()  As moddat_tpo_Genera

Private Sub cmb_Buscar_Click()
    If (cmb_Buscar.ListIndex = 0 Or cmb_Buscar.ListIndex = -1) Then
        txt_Buscar.Enabled = False
        Call gs_SetFocus(cmd_Buscar)
    Else
        txt_Buscar.Enabled = True
    End If
    txt_Buscar.Text = ""
End Sub

Private Sub cmb_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_Buscar.Enabled = False) Then
          Call gs_SetFocus(cmd_Buscar)
      Else
          Call gs_SetFocus(txt_Buscar)
      End If
   End If
End Sub


Private Sub cmb_Empres_Click()
   If cmb_Empres.ListIndex > -1 Then
      Call gs_SetFocus(cmb_Sucurs)
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmb_Sucurs_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Sucurs_Click
   End If
End Sub

Private Sub cmd_Borrar_Click()
Dim r_str_NumCta_Ref   As String

   moddat_g_str_Codigo = ""
   moddat_g_int_Situac = 0
   moddat_g_str_Situac = ""
   r_str_NumCta_Ref = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 12
   If CInt(grd_Listad.Text) <> 1 And CInt(grd_Listad.Text) <> 2 Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "Solo se pueden eliminar registros con estado vigente y vencido.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_Listad.Col = 14
   If grd_Listad.Text <> "" Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se puede eliminar el registro esta con origen renovación.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
   If MsgBox("¿Seguro que desea eliminar el registro seleccionado?" & vbCrLf & _
             "Recuerde que debe eliminar el asiento contable manual.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CLng(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
   
   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_MAEDPF_BORRAR ( "
   g_str_Parame = g_str_Parame & "'" & CLng(moddat_g_str_Codigo) & "', " 'MAEDPF_NUMCTA
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo completar la eliminación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   Else
      MsgBox "El registro se elimino, recuerde que debe eliminar el asiento contable manual.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
   
   Call fs_BuscarComp
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Buscar_Click()
   Call fs_BuscarComp
   cmb_Empres.Enabled = False
   cmb_Sucurs.Enabled = False
   ipp_FecIni.Enabled = False
   ipp_FecFin.Enabled = False
   chk_Estado.Enabled = False
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_Situac = 0
   moddat_g_str_Situac = ""
   
   moddat_g_str_Situac = "VIGENTE" 'Descripcion
   moddat_g_int_Situac = 1 'Codigo
   moddat_g_int_FlgGrb = 1
   frm_Ctb_InvDpf_02.Show 1
End Sub

Private Sub cmd_Consul_Click()
   moddat_g_str_Codigo = ""
   moddat_g_int_Situac = 0
   moddat_g_str_Situac = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   Call gs_RefrescaGrid(grd_Listad)
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   grd_Listad.Col = 13
   moddat_g_str_Situac = CStr(grd_Listad.Text) 'Descripcion
   grd_Listad.Col = 12
   moddat_g_int_Situac = CInt(grd_Listad.Text) 'Codigo
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 0
   frm_Ctb_InvDpf_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Detalle_Click()
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 0
   frm_Ctb_InvDpf_03.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_str_Codigo = ""
   moddat_g_int_Situac = 0
   moddat_g_str_Situac = ""
   moddat_g_int_TipObs = 0 'Tipo Deposito Plazo Fijo
   moddat_g_str_Observ = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   grd_Listad.Col = 12
   If CInt(grd_Listad.Text) <> 1 Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se pudo editar el registro, solo se editan con estado vigente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
 
   Call gs_RefrescaGrid(grd_Listad)
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   grd_Listad.Col = 13
   moddat_g_str_Situac = CStr(grd_Listad.Text) 'Descripcion
   grd_Listad.Col = 12
   moddat_g_int_Situac = CInt(grd_Listad.Text) 'Codigo
   
   moddat_g_int_FlgGrb = 2
   frm_Ctb_InvDpf_02.Show 1
   
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

Private Sub cmd_Generar_Click()
Dim r_int_Contad As Integer
Dim r_bol_Estado As Boolean

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If chk_Estado.Value = 0 Then
      MsgBox "Para poder generar el proceso de devengados, se tiene que activar la opción Todos los Estados.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_Ctb_InvDpf_05.Show 1
End Sub

Private Sub cmd_Gestion_Click()
   moddat_g_str_Codigo = ""
   moddat_g_int_Situac = 0
   moddat_g_str_Situac = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   grd_Listad.Col = 12
   'VIGENTE, VENCIDO
   If CInt(grd_Listad.Text) = 3 Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se puede gestionar registros con estado cerrado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   grd_Listad.Col = 13
   moddat_g_str_Situac = CStr(grd_Listad.Text) 'Descripcion
   grd_Listad.Col = 12
   moddat_g_int_Situac = CInt(grd_Listad.Text) 'Codigo
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 3
   frm_Ctb_InvDpf_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   cmb_Empres.Enabled = True
   cmb_Sucurs.Enabled = True
   ipp_FecIni.Enabled = True
   ipp_FecFin.Enabled = True
   chk_Estado.Enabled = True
   Call gs_SetFocus(cmb_Empres)
End Sub

Private Sub cmd_Reversa_Click()
   moddat_g_str_Codigo = ""
   moddat_g_int_Situac = 0
   moddat_g_str_Situac = ""
   moddat_g_int_TipObs = 0 'Tipo Deposito Plazo Fijo
   moddat_g_str_Observ = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   grd_Listad.Col = 13
   moddat_g_str_Situac = CStr(grd_Listad.Text) 'Descripcion
   grd_Listad.Col = 12
   moddat_g_int_Situac = CInt(grd_Listad.Text) 'Codigo
   grd_Listad.Col = 18
   moddat_g_int_TipObs = CInt(grd_Listad.Text) 'MAEDPF_TIPDPF
   grd_Listad.Col = 14
   moddat_g_str_Observ = CStr(grd_Listad.Text) 'MAEDPF_NUMCTA_REF
   Call gs_RefrescaGrid(grd_Listad)
 
   If moddat_g_int_TipObs <> 1 Or Trim(moddat_g_str_Observ) = "" Then
      If moddat_g_int_TipObs = 2 And Trim(moddat_g_str_Observ) = "" Then
      Else
         MsgBox "No se pueden reversar: " & Chr(13) & "----------------------------" & Chr(13) & "1. Los registros que hayan sido renovados." & Chr(13) & "2. Los registros vigentes y vencidos que no tengan dependencia.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If

   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 4
   frm_Ctb_InvDpf_02.Show 1
   Call gs_SetFocus(grd_Listad)
   
'   If moddat_g_int_TipObs <> 1 Or Trim(moddat_g_str_Observ) = "" Then
'      MsgBox "Solo se pueden revertir los registros que tienen renovación.", vbExclamation, modgen_g_str_NomPlt
'      Exit Sub
'   End If
'
'   If moddat_g_int_TipObs = 1 And CLng(Trim(moddat_g_str_Observ)) > 0 Then
'      grd_Listad.Col = 12
'      If CInt(grd_Listad.Text) <> 1 Then
'         Call gs_RefrescaGrid(grd_Listad)
'         MsgBox "Solo se pueden reviertir los registros con estado vigente.", vbExclamation, modgen_g_str_NomPlt
'         Exit Sub
'      End If
'
'      Call gs_RefrescaGrid(grd_Listad)
'      moddat_g_int_FlgGrb = 4
'      frm_Ctb_InvDpf_02.Show 1
'
'      Call gs_SetFocus(grd_Listad)
'   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()

   Call moddat_gs_FecSis
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "NINGUNA"
   cmb_Buscar.AddItem "INSTITUCION"
   cmb_Buscar.AddItem "OPERACION REF."
   cmb_Buscar.AddItem "ESTADO"
   
   grd_Listad.ColWidth(0) = 1020 'NRO CUENTA
   grd_Listad.ColWidth(1) = 2490 'INSITUCION
   grd_Listad.ColWidth(2) = 1800 'OPERACION DE REF.
   grd_Listad.ColWidth(3) = 870 'PLAZO
   grd_Listad.ColWidth(4) = 780 'TASA
   grd_Listad.ColWidth(5) = 800 'MONEDA
   grd_Listad.ColWidth(6) = 1380 'CAPITAL
   grd_Listad.ColWidth(7) = 1120 'FECHA APERTURA
   grd_Listad.ColWidth(8) = 1110 'FECHA VECIMIENTO
   grd_Listad.ColWidth(9) = 1080 'RENDIMIENTO
   grd_Listad.ColWidth(10) = 1140 'DEVENGADO
   grd_Listad.ColWidth(11) = 840 'POR VENCER DIAS
   grd_Listad.ColWidth(12) = 0 'COD_SITUACION
   grd_Listad.ColWidth(13) = 1040 'NOM_SITUACION
   grd_Listad.ColWidth(14) = 0 'MAEDPF_NUMCTA_REF
   grd_Listad.ColWidth(15) = 0 'CTADPF_CODENT_DES
   grd_Listad.ColWidth(16) = 0 'CTADPF_CODENT_ORI
   grd_Listad.ColWidth(17) = 0 'CTADPF_CODMON
   grd_Listad.ColWidth(18) = 0 'MAEDPF_TIPDPF
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   grd_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
   grd_Listad.ColAlignment(13) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
Dim r_str_CadAux As String

   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0
   cmb_Empres.ListIndex = 0
   r_str_CadAux = ""
   
   Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   
   pnl_Period.Caption = moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   r_str_CadAux = DateAdd("m", 1, "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno)
   modctb_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   modctb_str_FecIni = DateAdd("m", -1, modctb_str_FecFin)
   modctb_str_FecIni = "01/" & Format(Month(modctb_str_FecIni), "00") & "/" & Year(modctb_str_FecIni)
   
   ipp_FecIni.Text = "01/01/2019" 'modctb_str_FecIni
   ipp_FecFin.Text = modctb_str_FecFin
   
   cmb_Buscar.ListIndex = 0
   cmb_Sucurs.ListIndex = 0
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Public Sub fs_BuscarComp()
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_str_Cadena     As String
Dim r_str_FecVct     As String
Dim r_str_FecApe     As String
Dim r_int_FecDif     As Integer
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   r_str_FecIni = Format(ipp_FecIni.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFin.Text, "yyyymmdd")
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT LPAD(A.MAEDPF_NUMCTA,8,'0') MAEDPF_NUMCTA, TRIM(B.PARDES_DESCRI) AS ENTIDAD_DEST, A.MAEDPF_NUMREF, A.MAEDPF_PLADIA,  "
   g_str_Parame = g_str_Parame & "       A.MAEDPF_TASINT, A.MAEDPF_CODMON, A.MAEDPF_SALCAP, A.MAEDPF_INTAJU,  "
   g_str_Parame = g_str_Parame & "       A.MAEDPF_INTCAP, TRIM(C.PARDES_DESCRI) AS ENTIDAD_ORIG, TRIM(D.PARDES_DESCRI) AS MONEDA,  "
   g_str_Parame = g_str_Parame & "       A.MAEDPF_FECAPE, TO_CHAR(TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA,'yyyymmdd') AS FEC_VCTO,  "
   g_str_Parame = g_str_Parame & "       DECODE(A.MAEDPF_TIPDPF,1,  "
   g_str_Parame = g_str_Parame & "              CASE  "
   'g_str_Parame = g_str_Parame & "                WHEN (TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA)-1 = TO_DATE(SYSDATE,'DD/MM/YY') THEN 'POR VENCER'  "
   g_str_Parame = g_str_Parame & "                WHEN (TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA) <= TO_DATE(SYSDATE,'DD/MM/YY') THEN 'VENCIDO'  "
   g_str_Parame = g_str_Parame & "                WHEN TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA > TO_DATE(SYSDATE,'DD/MM/YY') THEN 'VIGENTE'  "
   g_str_Parame = g_str_Parame & "              END, 'CERRADO') AS NOM_SITUAC,  "
   g_str_Parame = g_str_Parame & "       DECODE(A.MAEDPF_TIPDPF,1,  "
   g_str_Parame = g_str_Parame & "       CASE  "
   g_str_Parame = g_str_Parame & "         WHEN (TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA) <= TO_DATE(SYSDATE,'DD/MM/YY') THEN 2  "
   g_str_Parame = g_str_Parame & "         WHEN TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA > TO_DATE(SYSDATE,'DD/MM/YY') THEN 1  "
   g_str_Parame = g_str_Parame & "       END, A.MAEDPF_SITDPF) AS COD_SITUAC, A.MAEDPF_NUMCTA_REF, MAEDPF_CODENT_DES, MAEDPF_CODENT_ORI, MAEDPF_TIPDPF  "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_MAEDPF A "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES B ON A.MAEDPF_CODENT_DES = B.PARDES_CODITE AND B.PARDES_CODGRP = 122  "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES C ON A.MAEDPF_CODENT_ORI = C.PARDES_CODITE AND C.PARDES_CODGRP = 122  "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON A.MAEDPF_CODMON = D.PARDES_CODITE AND D.PARDES_CODGRP = 204  "
   g_str_Parame = g_str_Parame & " WHERE A.MAEDPF_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "   AND A.MAEDPF_FECAPE BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   
   If (cmb_Buscar.ListIndex = 1) Then 'INSTITUCION
       If Len(Trim(txt_Buscar.Text)) > 0 Then
          g_str_Parame = g_str_Parame & "   AND TRIM(B.PARDES_DESCRI) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'"
       End If
   ElseIf (cmb_Buscar.ListIndex = 2) Then 'OPERACION REFERENCIA
       If Len(Trim(txt_Buscar.Text)) > 0 Then
          g_str_Parame = g_str_Parame & "   AND A.MAEDPF_NUMREF LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'"
       End If
   ElseIf (cmb_Buscar.ListIndex = 3) Then 'ESTADO
       If Len(Trim(txt_Buscar.Text)) > 0 Then
          g_str_Parame = g_str_Parame & "   AND (DECODE(A.MAEDPF_TIPDPF,1,  "
          g_str_Parame = g_str_Parame & "               CASE  "
          g_str_Parame = g_str_Parame & "                 WHEN (TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA) <= TO_DATE(SYSDATE,'DD/MM/YY') THEN 'VENCIDO'  "
          g_str_Parame = g_str_Parame & "                 WHEN TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA > TO_DATE(SYSDATE,'DD/MM/YY') THEN 'VIGENTE'  "
          g_str_Parame = g_str_Parame & "               END, 'CERRADO')) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'  "
       End If
   End If
   If chk_Estado.Value = False Then
      g_str_Parame = g_str_Parame & "   AND A.MAEDPF_SITDPF <> 3  "
   End If
   g_str_Parame = g_str_Parame & " ORDER BY A.MAEDPF_NUMCTA ASC "

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
      grd_Listad.Text = CStr(g_rst_Princi!MAEDPF_NUMCTA)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!ENTIDAD_DEST & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!MAEDPF_NUMREF & "")
      
      grd_Listad.Col = 3
      grd_Listad.Text = g_rst_Princi!MAEDPF_PLADIA
      
      grd_Listad.Col = 4
      grd_Listad.Text = Format(Trim(g_rst_Princi!MAEDPF_TASINT & ""), "###,###,###,##0.00")
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!Moneda & "")
                                    
      grd_Listad.Col = 6
      grd_Listad.Text = Format(g_rst_Princi!MAEDPF_SALCAP, "###,###,###,##0.00")

      grd_Listad.Col = 7
      grd_Listad.Text = gf_FormatoFecha(Trim(g_rst_Princi!MAEDPF_FECAPE & ""))
      r_str_FecApe = grd_Listad.Text
      
      grd_Listad.Col = 8
      grd_Listad.Text = DateAdd("d", g_rst_Princi!MAEDPF_PLADIA, gf_FormatoFecha(Trim(g_rst_Princi!MAEDPF_FECAPE & "")))
      r_str_FecVct = grd_Listad.Text

      grd_Listad.Col = 9
      grd_Listad.Text = Format(CDbl(g_rst_Princi!MAEDPF_INTCAP), "###,###,###,##0.00")
                        
      '=SI(F_DIA=F_VCTO,INTCAP, SI(F_DIA>F_VCTO,INTCAP, ((((1+TASA)^(1/360))-1)*CAPITAL)*(F_DIA-F_APER+1) ))
      grd_Listad.Col = 10
      grd_Listad.Text = "0.00"
      r_str_FecApe = Format(r_str_FecApe, "yyyymmdd")
      r_str_FecVct = Format(r_str_FecVct, "yyyymmdd")
      r_int_FecDif = 0

      Dim ImpAux As String
      Call fs_CalDevengado(moddat_g_str_FecSis, r_str_FecVct, CDbl(g_rst_Princi!MAEDPF_INTCAP), _
                           r_str_FecApe, g_rst_Princi!MAEDPF_TASINT, g_rst_Princi!MAEDPF_SALCAP, ImpAux)
      grd_Listad.Text = ImpAux
                           
      If g_rst_Princi!COD_SITUAC <> 3 Then
         grd_Listad.Col = 11
         grd_Listad.Text = DateDiff("d", moddat_g_str_FecSis, gf_FormatoFecha(r_str_FecVct))
      End If
      
      grd_Listad.Col = 12
      grd_Listad.Text = Trim(g_rst_Princi!COD_SITUAC & "")
      
      grd_Listad.Col = 13
      grd_Listad.Text = Trim(g_rst_Princi!NOM_SITUAC & "")
      
      grd_Listad.Col = 14
      grd_Listad.Text = Trim(g_rst_Princi!MAEDPF_NUMCTA_REF & "")
      
      grd_Listad.Col = 15
      grd_Listad.Text = Trim(g_rst_Princi!MAEDPF_CODENT_DES & "")
      
      grd_Listad.Col = 16
      grd_Listad.Text = Trim(g_rst_Princi!MAEDPF_CODENT_ORI & "")
      
      grd_Listad.Col = 17
      grd_Listad.Text = Trim(g_rst_Princi!MAEDPF_CODMON & "")
      
      grd_Listad.Col = 18
      grd_Listad.Text = Trim(g_rst_Princi!MAEDPF_TIPDPF & "")
            
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Public Function fs_CalDevengado(p_FecAct As String, p_FecVct As String, p_IntCap As Double, p_FecApe As String, p_TasInt As Double, p_SalCap As Double, ByRef p_Importe As String)
Dim r_int_FecDif   As Integer
   
   r_int_FecDif = 0
   p_Importe = "0"
   
   If Format(p_FecAct, "yyyymmdd") = p_FecVct Then
      p_Importe = Format(p_IntCap, "###,###,###,##0.00")
   ElseIf Format(p_FecAct, "yyyymmdd") > p_FecVct Then
      p_Importe = Format(p_IntCap, "###,###,###,##0.00")
   Else
      r_int_FecDif = DateDiff("d", gf_FormatoFecha(p_FecApe), p_FecAct) + 1
      p_Importe = CStr(((((1 + (p_TasInt / 100)) ^ (1 / 360)) - 1) * p_SalCap) * r_int_FecDif)
      p_Importe = Format(p_Importe, "###,###,###,##0.00")
   End If
      
End Function

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contar        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE INVERSIONES - DEPOSITO PLAZO FIJO"
      .Range(.Cells(2, 2), .Cells(2, 13)).Merge
      .Range(.Cells(2, 2), .Cells(2, 13)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 13)).HorizontalAlignment = xlHAlignCenter

      .Cells(3, 2) = "NRO CUENTA"
      .Cells(3, 3) = "INSTITUCION"
      .Cells(3, 4) = "OPERACIONE DE REF."
      .Cells(3, 5) = "PLAZO DIAS"
      .Cells(3, 6) = "TASA %"
      .Cells(3, 7) = "MONEDA"
      .Cells(3, 8) = "CAPITAL"
      .Cells(3, 9) = "F.APERTURA"
      .Cells(3, 10) = "F.VENCIMIENTO"
      .Cells(3, 11) = "RENDIMIENTO"
      .Cells(3, 12) = "DEVENGADO"
      .Cells(3, 13) = "ESTADO"
         
      .Range(.Cells(3, 2), .Cells(3, 13)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 13)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13 'Nro Cuenta
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 35 'institucion
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 20 'Operacion de Ref.
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 11 'Plazo Dias
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 11 'Tasa %
      .Columns("F").NumberFormat = "#,###0.00"
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 22 'Moneda
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 18 'Capital
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 14 'F.Apertura
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("I").NumberFormat = "dd-mm-yyyy"
      .Columns("J").ColumnWidth = 16 'F.Vencimiento
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("J").NumberFormat = "dd-mm-yyyy"
      .Columns("K").ColumnWidth = 16 'Rendimiento
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 16 'Devengado
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 15 'Estado
      .Columns("M").HorizontalAlignment = xlHAlignCenter
            
      .Range(.Cells(1, 1), .Cells(10, 13)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 13)).Font.Size = 11
      
      r_int_NumFil = 2
      For r_int_Contar = 0 To grd_Listad.Rows - 1
          .Cells(r_int_NumFil + 2, 2) = "'" & grd_Listad.TextMatrix(r_int_Contar, 0) 'Nro Cuenta
          .Cells(r_int_NumFil + 2, 3) = "'" & grd_Listad.TextMatrix(r_int_Contar, 1) 'institucion
          .Cells(r_int_NumFil + 2, 4) = "'" & grd_Listad.TextMatrix(r_int_Contar, 2) 'Operacion de Ref.
          .Cells(r_int_NumFil + 2, 5) = grd_Listad.TextMatrix(r_int_Contar, 3)  'Plazo Dias
          .Cells(r_int_NumFil + 2, 6) = grd_Listad.TextMatrix(r_int_Contar, 4) 'Tasa %
          .Cells(r_int_NumFil + 2, 7) = "'" & grd_Listad.TextMatrix(r_int_Contar, 5) 'Moneda
          .Cells(r_int_NumFil + 2, 8) = grd_Listad.TextMatrix(r_int_Contar, 6)  'Capital
          .Cells(r_int_NumFil + 2, 9) = CDate(grd_Listad.TextMatrix(r_int_Contar, 7))  'F.Apertura
          .Cells(r_int_NumFil + 2, 10) = CDate(grd_Listad.TextMatrix(r_int_Contar, 8))   'F.Vencimiento
          .Cells(r_int_NumFil + 2, 11) = grd_Listad.TextMatrix(r_int_Contar, 9) 'Rendimiento
          .Cells(r_int_NumFil + 2, 12) = grd_Listad.TextMatrix(r_int_Contar, 10)  'Devengado
          .Cells(r_int_NumFil + 2, 13) = "'" & grd_Listad.TextMatrix(r_int_Contar, 13) 'Estado
          
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(3, 3), .Cells(3, 13)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(chk_Estado)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub chk_Estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub txt_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call fs_BuscarComp
   Else
      If (cmb_Buscar.ListIndex = 1) Then
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
      Else
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
      End If
   End If
End Sub
