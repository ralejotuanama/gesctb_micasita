VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_Provis_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   9915
   ClientLeft      =   6675
   ClientTop       =   1845
   ClientWidth     =   9030
   Icon            =   "GesCtb_frm_832.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9045
      _Version        =   65536
      _ExtentX        =   15954
      _ExtentY        =   17806
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
         Height          =   2205
         Left            =   30
         TabIndex        =   15
         Top             =   7680
         Width           =   8955
         _Version        =   65536
         _ExtentX        =   15796
         _ExtentY        =   3889
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
            Height          =   1785
            Left            =   60
            TabIndex        =   11
            Top             =   360
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   3149
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_CodEmp 
            Height          =   285
            Left            =   90
            TabIndex        =   25
            Top             =   60
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   1830
            TabIndex        =   26
            Top             =   60
            Width           =   7035
            _Version        =   65536
            _ExtentX        =   12409
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre"
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   8955
         _Version        =   65536
         _ExtentX        =   15796
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
            TabIndex        =   17
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
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
            Picture         =   "GesCtb_frm_832.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1845
         Left            =   30
         TabIndex        =   18
         Top             =   5790
         Width           =   8955
         _Version        =   65536
         _ExtentX        =   15796
         _ExtentY        =   3254
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
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   7740
            Picture         =   "GesCtb_frm_832.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Buscar Clientes por Apellidos y Nombres"
            Top             =   30
            Width           =   585
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1050
            Width           =   2775
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1890
            MaxLength       =   30
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   1380
            Width           =   2775
         End
         Begin VB.CommandButton cmd_LimBus 
            Height          =   585
            Left            =   8340
            Picture         =   "GesCtb_frm_832.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
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
            Left            =   120
            TabIndex        =   31
            Top             =   180
            Width           =   2985
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   60
            TabIndex        =   20
            Top             =   1050
            Width           =   1725
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   60
            TabIndex        =   19
            Top             =   1380
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1485
         Left            =   30
         TabIndex        =   22
         Top             =   1830
         Width           =   8955
         _Version        =   65536
         _ExtentX        =   15796
         _ExtentY        =   2619
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
            Left            =   7740
            Picture         =   "GesCtb_frm_832.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Buscar Operaciones por Documento de Identidad"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   8340
            Picture         =   "GesCtb_frm_832.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   930
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   600
            Width           =   2775
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
            TabIndex        =   30
            Top             =   60
            Width           =   3885
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   24
            Top             =   930
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   600
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   1035
         Left            =   30
         TabIndex        =   27
         Top             =   750
         Width           =   8955
         _Version        =   65536
         _ExtentX        =   15796
         _ExtentY        =   1826
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
            Left            =   8340
            Picture         =   "GesCtb_frm_832.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmb_LimOpe 
            Height          =   585
            Left            =   7740
            Picture         =   "GesCtb_frm_832.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar todas las Búsquedas"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   7140
            Picture         =   "GesCtb_frm_832.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Crédito por Número de Operación"
            Top             =   30
            Width           =   585
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1890
            TabIndex        =   1
            Top             =   540
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
            Caption         =   "Búsqueda por Número de Operación"
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
            TabIndex        =   29
            Top             =   90
            Width           =   3885
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Operación:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   540
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2400
         Left            =   30
         TabIndex        =   32
         Top             =   3360
         Width           =   8955
         _Version        =   65536
         _ExtentX        =   15796
         _ExtentY        =   4233
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
            Height          =   1965
            Left            =   60
            TabIndex        =   33
            Top             =   390
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   3466
            _Version        =   393216
            Rows            =   21
            Cols            =   3
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
            TabIndex        =   34
            Top             =   90
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
            TabIndex        =   35
            Top             =   90
            Width           =   3810
            _Version        =   65536
            _ExtentX        =   6720
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
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
            Left            =   5880
            TabIndex        =   36
            Top             =   90
            Width           =   2985
            _Version        =   65536
            _ExtentX        =   5265
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Clasificación de Provisión"
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
Attribute VB_Name = "frm_Mnt_Provis_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_BusPer_Click()
   Dim r_int_PerMes  As Integer
   Dim r_int_PerAno  As Integer
  
   g_str_Parame = "SELECT * FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_PerMes = g_rst_Princi!HIPCIE_PERMES
   r_int_PerAno = g_rst_Princi!HIPCIE_PERANO
   
   modsec_g_str_Period = r_int_PerAno & Format(r_int_PerMes, "00")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
End Sub


Private Sub cmd_Buscar_Click()
   If Len(Trim(msk_NumOpe.Text)) < 10 Then
      MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(msk_NumOpe)
      Exit Sub
   End If
   
   txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
   
   'moddat_g_str_NumOpe = Left(msk_NumOpe.Text, 3) & Mid(msk_NumOpe.Text, 5, 2) & Right(msk_NumOpe.Text, 5)
   moddat_g_str_NumOpe = msk_NumOpe.Text
   
   g_str_Parame = "SELECT * FROM CRE_HIPCIE WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado clientes para esa selección.", vbExclamation, modgen_g_str_NomPlt
      
   Else
      moddat_g_int_TipDoc = Trim(g_rst_Princi!HIPCIE_TDOCLI)
      moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPCIE_NDOCLI)
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      frm_Mnt_Provis_04.Show 1
   End If

End Sub

Private Sub cmb_LimOpe_Click()
   msk_NumOpe.Mask = ""
   msk_NumOpe.Text = ""
   msk_NumOpe.Mask = "###-##-#####"
   
   Call cmd_Limpia_Click
   Call cmd_LimBus_Click
   
   Call gs_SetFocus(msk_NumOpe)
End Sub


Private Sub cmd_BusCli_Click()
   Dim r_str_ApePat  As String
   Dim r_str_ApeMat  As String
   Dim r_str_Nombre  As String

   If Len(Trim(txt_ApePat)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_con_AteCli
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
   
   r_str_ApePat = txt_ApePat.Text & "%"
   r_str_ApeMat = txt_ApeMat.Text & "%"
   r_str_Nombre = txt_Nombre.Text & "%"
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "RTRIM(DATGEN_APEPAT) LIKE '" & r_str_ApePat & "' AND "
   g_str_Parame = g_str_Parame & "RTRIM(DATGEN_APEMAT) LIKE '" & r_str_ApeMat & "' AND "
   g_str_Parame = g_str_Parame & "RTRIM(DATGEN_NOMBRE) LIKE '" & r_str_Nombre & "' ORDER BY "
   g_str_Parame = g_str_Parame & "DATGEN_APEPAT ASC, "
   g_str_Parame = g_str_Parame & "DATGEN_APEMAT ASC, "
   g_str_Parame = g_str_Parame & "DATGEN_NOMBRE ASC "
   
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se han encontrado clientes para esta selección.", vbExclamation, modgen_g_con_AteCli
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!DATGEN_TIPDOC) & "-" & Trim(g_rst_Princi!DATGEN_NUMDOC & "")
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!DatGen_ApePat & "") & " " & Trim(g_rst_Princi!DatGen_ApeMat & "") & " " & Trim(g_rst_Princi!DatGen_Nombre & "")
      
      g_rst_Princi.MoveNext
   Loop
         
   grd_Listad.Redraw = True
   
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_LimBus_Click()
   Call fs_Limpia_BusAlf
   
   Call gs_SetFocus(txt_ApePat)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Call cmd_Limpia_Click
   Call fs_Limpia_BusAlf
   Call cmd_Limpia_Click
   Call cmb_LimOpe_Click
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call gs_CentraForm(Me)
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicio
   Call cmd_BusPer_Click
   Call cmd_Limpia_Click
   Call fs_Limpia_BusAlf
   Call cmd_Limpia_Click
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   grd_Listad.ColWidth(0) = 1800
   grd_Listad.ColWidth(1) = 7000
      
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   grd_LisOpe.ColWidth(0) = 2000
   grd_LisOpe.ColWidth(1) = 3800
   grd_LisOpe.ColWidth(2) = 3000
   
   grd_LisOpe.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(1) = flexAlignLeftCenter
   grd_LisOpe.ColAlignment(2) = flexAlignCenterCenter
   
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_LimpiaGrid(grd_LisOpe)
         
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
End Sub

Private Sub fs_Limpia_BusAlf()
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_Nombre.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub cmd_Limpia_Click()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
   Call gs_SetFocus(cmb_TipDoc)
   Call gs_LimpiaGrid(grd_LisOpe)
End Sub

Private Sub grd_LisOpe_DblClick()
   Dim r_str_NumOpe     As String

   If grd_LisOpe.Rows = 0 Then
      Exit Sub
   End If
   
   grd_LisOpe.Col = 0
   'r_str_NumOpe = Left(grd_LisOpe.Text, 3) & Mid(grd_LisOpe.Text, 5, 2) & Right(grd_LisOpe.Text, 5)
   r_str_NumOpe = grd_LisOpe.Text
   
   Call gs_RefrescaGrid(grd_LisOpe)
   
   msk_NumOpe.Text = r_str_NumOpe
   Call cmd_Buscar_Click
End Sub


Private Sub grd_Listad_DblClick()
   Dim r_int_TipDoc     As Integer
   Dim r_str_NumDoc     As String

   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 0
      
      moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
      moddat_g_str_NumDoc = Right(grd_Listad.Text, 8)
      
      r_int_TipDoc = CInt(Left(grd_Listad.Text, 1))
      r_str_NumDoc = Mid(grd_Listad.Text, 3)
   
      Call gs_RefrescaGrid(grd_Listad)
      
      Call gs_BuscarCombo_Item(cmb_TipDoc, r_int_TipDoc)
      txt_NumDoc.Text = r_str_NumDoc
      
      Call cmd_BusOpe_Click
      Call gs_SetFocus(grd_LisOpe)
   End If
   
   'If grd_LisOpe.Rows > 0 Then
   '   grd_LisOpe.Col = 0
   
   '   modsec_g_str_NumOpe = grd_LisOpe.Text
   'End If

End Sub

Private Sub msk_NumOpe_GotFocus()
   Call gs_SelecTodo(msk_NumOpe)
End Sub

Private Sub msk_NumOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
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
            Case 4:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub cmd_BusOpe_Click()
   Dim r_int_FlgEnc  As Integer

   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   r_int_FlgEnc = 0
   
   grd_LisOpe.Redraw = False
   
   Call gs_LimpiaGrid(grd_LisOpe)
   
   'Buscando Operaciones como Cliente Titular
   g_str_Parame = "SELECT * FROM CRE_HIPCIE WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_TDOCLI = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_NDOCLI = '" & txt_NumDoc.Text & "' AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & Right(modsec_g_str_Period, 2) & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & Left(modsec_g_str_Period, 4) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgEnc = 1
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_LisOpe.Rows = grd_LisOpe.Rows + 1
         grd_LisOpe.Row = grd_LisOpe.Rows - 1
         
         grd_LisOpe.Col = 0
         grd_LisOpe.Text = Mid(g_rst_Princi!HIPCIE_NUMOPE, 1, 3) & "-" & Mid(g_rst_Princi!HIPCIE_NUMOPE, 4, 2) & "-" & Mid(g_rst_Princi!HIPCIE_NUMOPE, 6, 5)
         
         grd_LisOpe.Col = 1
         grd_LisOpe.Text = moddat_gf_Consulta_Produc(g_rst_Princi!HIPCIE_CODPRD)
         
         grd_LisOpe.Col = 2
         grd_LisOpe.Text = modsec_gf_Buscar_TipCla(4, g_rst_Princi!HIPCIE_CLAPRV)
         
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Como Cónyuge
   'g_str_Parame = "SELECT * FROM CRE_HIPCIE WHERE "
   'g_str_Parame = g_str_Parame & "HIPCIE_TDOCYG = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
   'g_str_Parame = g_str_Parame & "HIPCIE_NDOCYG = '" & txt_NumDoc.Text & "' "
   
   'If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
   '   Exit Sub
   'End If
   
   'If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   '   r_int_FlgEnc = 1
   '
   '   g_rst_Princi.MoveFirst
   '
   '   Do While Not g_rst_Princi.EOF
   '      grd_LisOpe.Rows = grd_LisOpe.Rows + 1
   '      grd_LisOpe.Row = grd_LisOpe.Rows - 1
   '
   '      grd_LisOpe.Col = 0
   '      grd_LisOpe.Text = Mid(g_rst_Princi!HIPCIE_NUMOPE, 1, 3) & "-" & Mid(g_rst_Princi!HIPCIE_NUMOPE, 4, 2) & "-" & Mid(g_rst_Princi!HIPCIE_NUMOPE, 6, 5)
   '
   '      grd_LisOpe.Col = 1
   '      grd_LisOpe.Text = moddat_gf_Consulta_Produc(g_rst_Princi!HIPCIE_CODPRD)
   '
   '      g_rst_Princi.MoveNext
   '   Loop
   'End If
   
   'g_rst_Princi.Close
   'Set g_rst_Princi = Nothing
   
   grd_LisOpe.Redraw = True
   
   If grd_LisOpe.Rows > 0 Then
      'Call pnl_Tit_NumOpe_Click
      
      Call gs_UbiIniGrid(grd_LisOpe)
   Else
      MsgBox "No se encontró ningún Crédito para este Documento de Identidad.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub


