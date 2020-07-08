VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Con_OpeFin_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10155
   ClientLeft      =   3060
   ClientTop       =   735
   ClientWidth     =   11910
   Icon            =   "GesCtb_frm_106.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10155
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11895
      _Version        =   65536
      _ExtentX        =   20981
      _ExtentY        =   17912
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   1095
         Left            =   30
         TabIndex        =   17
         Top             =   750
         Width           =   11805
         _Version        =   65536
         _ExtentX        =   20823
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11130
            Picture         =   "GesCtb_frm_106.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   390
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   10530
            Picture         =   "GesCtb_frm_106.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar todas las Búsquedas"
            Top             =   390
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   9930
            Picture         =   "GesCtb_frm_106.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Crédito por Número de Operación"
            Top             =   390
            Width           =   585
         End
         Begin VB.ComboBox cmb_SucAge 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   390
            Width           =   7185
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1890
            TabIndex        =   1
            Top             =   750
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "##-#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
            Caption         =   "Búsqueda por Número de Movimiento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   20
            Top             =   90
            Width           =   3885
         End
         Begin VB.Label Label9 
            Caption         =   "Sucursal:"
            Height          =   225
            Left            =   90
            TabIndex        =   19
            Top             =   390
            Width           =   945
         End
         Begin VB.Label Label10 
            Caption         =   "Nro. Movimiento:"
            Height          =   285
            Left            =   90
            TabIndex        =   18
            Top             =   720
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1095
         Left            =   30
         TabIndex        =   21
         Top             =   1890
         Width           =   11805
         _Version        =   65536
         _ExtentX        =   20823
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
         Begin VB.CommandButton cmd_BusOpe 
            Height          =   585
            Left            =   10530
            Picture         =   "GesCtb_frm_106.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Buscar Operaciones por Documento de Identidad"
            Top             =   360
            Width           =   585
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   390
            Width           =   2775
         End
         Begin VB.CommandButton cmd_LimOpe 
            Height          =   585
            Left            =   11130
            Picture         =   "GesCtb_frm_106.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Limpiar Lista de Coincidencias por Documento de Identidad"
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Identidad:"
            Height          =   285
            Left            =   90
            TabIndex        =   24
            Top             =   720
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   23
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Búsqueda por Documento de Identidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   90
            Width           =   3885
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1425
         Left            =   30
         TabIndex        =   25
         Top             =   6030
         Width           =   11805
         _Version        =   65536
         _ExtentX        =   20823
         _ExtentY        =   2514
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
         Begin VB.CommandButton cmd_LimCli 
            Height          =   585
            Left            =   11160
            Picture         =   "GesCtb_frm_106.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Limpiar Lista de Coincidencias por Apellidos y Nombres"
            Top             =   360
            Width           =   585
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   1050
            Width           =   2775
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   10560
            Picture         =   "GesCtb_frm_106.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Buscar Clientes por Apellidos y Nombres"
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   29
            Top             =   390
            Width           =   1725
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   27
            Top             =   1050
            Width           =   1725
         End
         Begin VB.Label Label7 
            Caption         =   "Búsqueda por Apellidos y Nombres"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   26
            Top             =   90
            Width           =   2985
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   30
         Top             =   30
         Width           =   11805
         _Version        =   65536
         _ExtentX        =   20823
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
            Height          =   255
            Left            =   630
            TabIndex        =   31
            Top             =   30
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Consulta de Operaciones Financieras "
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   255
            Left            =   630
            TabIndex        =   32
            Top             =   330
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Búsqueda por Número de Movimiento"
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
            Picture         =   "GesCtb_frm_106.frx":168A
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   2580
         Left            =   30
         TabIndex        =   33
         Top             =   7500
         Width           =   11805
         _Version        =   65536
         _ExtentX        =   20823
         _ExtentY        =   4551
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_LisCli 
            Height          =   2145
            Left            =   60
            TabIndex        =   15
            Top             =   390
            Width           =   11685
            _ExtentX        =   20611
            _ExtentY        =   3784
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   90
            TabIndex        =   34
            Top             =   90
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento Identidad"
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   2070
            TabIndex        =   35
            Top             =   90
            Width           =   9330
            _Version        =   65536
            _ExtentX        =   16457
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Cliente"
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   2970
         Left            =   30
         TabIndex        =   36
         Top             =   3030
         Width           =   11805
         _Version        =   65536
         _ExtentX        =   20823
         _ExtentY        =   5239
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_LisOpe 
            Height          =   2505
            Left            =   60
            TabIndex        =   9
            Top             =   390
            Width           =   11685
            _ExtentX        =   20611
            _ExtentY        =   4419
            _Version        =   393216
            Rows            =   21
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   90
            TabIndex        =   37
            Top             =   90
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Refer. (Ope. / Solic.)"
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   2070
            TabIndex        =   38
            Top             =   90
            Width           =   840
            _Version        =   65536
            _ExtentX        =   1482
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Sucursal"
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
            Left            =   2910
            TabIndex        =   39
            Top             =   90
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Movim"
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
            Left            =   4170
            TabIndex        =   40
            Top             =   90
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Movim"
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
            Left            =   5400
            TabIndex        =   41
            Top             =   90
            Width           =   3450
            _Version        =   65536
            _ExtentX        =   6085
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Movimiento"
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
            Left            =   8850
            TabIndex        =   42
            Top             =   90
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   1958
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   9960
            TabIndex        =   43
            Top             =   90
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
   End
End
Attribute VB_Name = "frm_Con_OpeFin_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_SucAge()      As moddat_tpo_Genera
Dim l_str_FecMov        As String

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   Dim r_str_NumMov     As String
   
   If Len(Trim(msk_NumOpe.Text)) < 7 Then
      MsgBox "Debe ingresar el Número de Movimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(msk_NumOpe)
      Exit Sub
   End If
   
   r_str_NumMov = msk_NumOpe.Text
   
   g_str_Parame = "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Codigo & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMMOV = " & Right(r_str_NumMov, 5) & " AND "
   
   If Len(Trim(l_str_FecMov)) > 0 Then
      g_str_Parame = g_str_Parame & "CAJMOV_FECMOV = " & l_str_FecMov & " "
   Else
      g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & "20" & Left(r_str_NumMov, 2) & "0101" & " AND "
      g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & "20" & Left(r_str_NumMov, 2) & "1231" & " "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      MsgBox "No existe ninguna Operación registrada con ese Número. ", vbExclamation, modgen_g_str_NomPlt
      
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   moddat_g_str_CodGrp = g_rst_Princi!CAJMOV_SUCMOV
   opecaj_g_str_NumMov = CStr(g_rst_Princi!CAJMOV_NUMMOV)
   opecaj_g_str_FecMov = CStr(g_rst_Princi!CAJMOV_FECMOV)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   frm_Con_OpeFin_02.Show 1
End Sub

Private Sub cmd_BusCli_Click()
   Dim r_str_ApePat  As String
   Dim r_str_ApeMat  As String
   Dim r_str_Nombre  As String

   If Len(Trim(txt_ApePat)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_LisCli)
   l_str_FecMov = ""
   
   r_str_ApePat = txt_ApePat.Text & "%"
   r_str_ApeMat = txt_ApeMat.Text & "%"
   r_str_Nombre = txt_Nombre.Text & "%"
   
   g_str_Parame = "SELECT * FROM CLI_BUSCLI WHERE "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_APEPAT) LIKE '" & r_str_ApePat & "' AND "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_APEMAT) LIKE '" & r_str_ApeMat & "' AND "
   g_str_Parame = g_str_Parame & "RTRIM(BUSCLI_NOMBRE) LIKE '" & r_str_Nombre & "' ORDER BY "
   g_str_Parame = g_str_Parame & "BUSCLI_APEPAT ASC, "
   g_str_Parame = g_str_Parame & "BUSCLI_APEMAT ASC, "
   g_str_Parame = g_str_Parame & "BUSCLI_NOMBRE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado clientes para esa selección.", vbExclamation, modgen_g_str_NomPlt
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_LisCli.Rows = grd_LisCli.Rows + 1
      
      grd_LisCli.Row = grd_LisCli.Rows - 1
      
      grd_LisCli.Col = 0
      grd_LisCli.Text = CStr(g_rst_Princi!BUSCLI_TIPDOC) & "-" & Trim(g_rst_Princi!BUSCLI_NUMDOC & "")
      
      grd_LisCli.Col = 1
      grd_LisCli.Text = Trim(g_rst_Princi!BUSCLI_APEPAT & "") & " " & Trim(g_rst_Princi!BUSCLI_APEMAT & "") & " " & Trim(g_rst_Princi!BUSCLI_NOMBRE & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   Call gs_UbiIniGrid(grd_LisCli)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_BusOpe_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   l_str_FecMov = ""
   
   Call gs_LimpiaGrid(grd_LisOpe)
   
   'Buscando Operaciones
   g_str_Parame = "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_TIPDOC = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMDOC = '" & txt_NumDoc.Text & "' "
   g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_FECMOV DESC, CAJMOV_NUMMOV DESC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_LisOpe.Redraw = False
   
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_LisOpe.Rows = grd_LisOpe.Rows + 1
         grd_LisOpe.Row = grd_LisOpe.Rows - 1
         
         grd_LisOpe.Col = 0
         
         If g_rst_Princi!CAJMOV_TIPMOV = 1101 Or g_rst_Princi!CAJMOV_TIPMOV = 2101 Then
            grd_LisOpe.Text = gf_Formato_NumSol(Trim(g_rst_Princi!CAJMOV_NUMOPE))
         Else
            grd_LisOpe.Text = gf_Formato_NumOpe(Trim(g_rst_Princi!CAJMOV_NUMOPE))
         End If
         
         grd_LisOpe.Col = 1
         grd_LisOpe.Text = Trim(g_rst_Princi!CAJMOV_SUCMOV)
         
         grd_LisOpe.Col = 2
         grd_LisOpe.Text = Mid(CStr(g_rst_Princi!CAJMOV_FECMOV), 3, 2) & "-" & Format(g_rst_Princi!CAJMOV_NUMMOV, "00000")
         
         grd_LisOpe.Col = 3
         grd_LisOpe.Text = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))
         
         grd_LisOpe.Col = 4
         grd_LisOpe.Text = CStr(g_rst_Princi!CAJMOV_TIPMOV) & " - " & moddat_gf_Consulta_ParDes("301", CStr(g_rst_Princi!CAJMOV_TIPMOV))
         
         grd_LisOpe.Col = 5
         grd_LisOpe.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!CAJMOV_MONPAG))
         
         grd_LisOpe.Col = 6
         grd_LisOpe.Text = Format(g_rst_Princi!CAJMOV_IMPTOT, "###,###,##0.00")
         
         g_rst_Princi.MoveNext
      Loop
      
      grd_LisOpe.Redraw = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_LisOpe.Rows > 0 Then
      Call gs_UbiIniGrid(grd_LisOpe)
   End If
End Sub

Private Sub cmd_LimCli_Click()
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   
   Call gs_LimpiaGrid(grd_LisCli)
End Sub

Private Sub cmd_LimOpe_Click()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
   Call gs_LimpiaGrid(grd_LisOpe)
End Sub

Private Sub cmd_Limpia_Click()
   msk_NumOpe.Mask = ""
   msk_NumOpe.Text = ""
   msk_NumOpe.Mask = "##-#####"
   
   Call cmd_LimOpe_Click
   Call cmd_LimCli_Click
   
   Call gs_SetFocus(msk_NumOpe)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicio
   
   Call gs_LimpiaGrid(grd_LisCli)
   Call gs_LimpiaGrid(grd_LisOpe)
   
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   'Cargando Listas de Sucursales
   moddat_g_str_Codigo = "000001"
   Call moddat_gs_Carga_SucAge(cmb_SucAge, l_arr_SucAge, moddat_g_str_Codigo)

   'Cargando Tipos de Documentos de Identidad
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")

   'Lista de Clientes
   grd_LisCli.ColWidth(0) = 2000
   grd_LisCli.ColWidth(1) = 9330
   
   grd_LisCli.ColAlignment(0) = flexAlignCenterCenter
   grd_LisCli.ColAlignment(1) = flexAlignLeftCenter
   
   'Lista de Movimientos
   grd_LisOpe.ColWidth(0) = 1995
   grd_LisOpe.ColWidth(1) = 840
   grd_LisOpe.ColWidth(2) = 1260
   grd_LisOpe.ColWidth(3) = 1230
   grd_LisOpe.ColWidth(4) = 3450
   grd_LisOpe.ColWidth(5) = 1110
   grd_LisOpe.ColWidth(6) = 1410
   grd_LisOpe.ColWidth(7) = 0
   grd_LisOpe.ColWidth(8) = 0
   grd_LisOpe.ColWidth(9) = 0
   
   grd_LisOpe.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(1) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(2) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(3) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(4) = flexAlignLeftCenter
   grd_LisOpe.ColAlignment(5) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(6) = flexAlignRightCenter
End Sub

Private Sub grd_LisCli_DblClick()
   Dim r_int_TipDoc     As Integer
   Dim r_str_NumDoc     As String

   If grd_LisCli.Rows > 0 Then
      grd_LisCli.Col = 0
      
      r_int_TipDoc = CInt(Left(grd_LisCli.Text, 1))
      r_str_NumDoc = Mid(grd_LisCli.Text, 3)
   
      Call gs_RefrescaGrid(grd_LisCli)
      
      Call gs_BuscarCombo_Item(cmb_TipDoc, r_int_TipDoc)
      txt_NumDoc.Text = r_str_NumDoc
      
      Call cmd_BusOpe_Click
      Call gs_SetFocus(grd_LisOpe)
   End If
End Sub

Private Sub grd_LisCli_SelChange()
   If grd_LisCli.Rows > 2 Then
      grd_LisCli.RowSel = grd_LisCli.Row
   End If
End Sub

Private Sub grd_LisOpe_DblClick()
   Dim r_str_NumOpe     As String

   If grd_LisOpe.Rows = 0 Then
      Exit Sub
   End If
   
   grd_LisOpe.Col = 1
   cmb_SucAge.ListIndex = gf_Busca_Arregl(l_arr_SucAge, grd_LisOpe.Text) - 1

   grd_LisOpe.Col = 2
   r_str_NumOpe = Left(grd_LisOpe.Text, 2) & Right(grd_LisOpe.Text, 5)
   
   grd_LisOpe.Col = 3
   l_str_FecMov = Format(CDate(grd_LisOpe.Text), "yyyymmdd")
   
   Call gs_RefrescaGrid(grd_LisOpe)
   
   msk_NumOpe.Text = r_str_NumOpe
   Call cmd_Buscar_Click
End Sub

Private Sub grd_LisOpe_SelChange()
   If grd_LisOpe.Rows > 2 Then
      grd_LisOpe.RowSel = grd_LisOpe.Row
   End If
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusCli)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " -_")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusOpe)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub


