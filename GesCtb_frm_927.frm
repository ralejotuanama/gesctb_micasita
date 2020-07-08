VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_AsiCtb_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   Icon            =   "GesCtb_frm_927.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7155
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9495
      _Version        =   65536
      _ExtentX        =   16748
      _ExtentY        =   12621
      _StockProps     =   15
      BackColor       =   14215660
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
         Height          =   4815
         Left            =   30
         TabIndex        =   8
         Top             =   2280
         Width           =   9405
         _Version        =   65536
         _ExtentX        =   16589
         _ExtentY        =   8493
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
         Begin Threed.SSPanel pnl_Tit_Codigo 
            Height          =   285
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   2220
            TabIndex        =   10
            Top             =   60
            Width           =   6795
            _Version        =   65536
            _ExtentX        =   11986
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripción Cuenta Contable"
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
            Height          =   4275
            Left            =   30
            TabIndex        =   2
            Top             =   360
            Width           =   9285
            _ExtentX        =   16378
            _ExtentY        =   7541
            _Version        =   393216
            Rows            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   11
         Top             =   60
         Width           =   9405
         _Version        =   65536
         _ExtentX        =   16589
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
            TabIndex        =   12
            Top             =   60
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Buscar Cuentas Contables"
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
            Picture         =   "GesCtb_frm_927.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   13
         Top             =   780
         Width           =   9405
         _Version        =   65536
         _ExtentX        =   16589
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_927.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Acepta 
            Height          =   585
            Left            =   1200
            Picture         =   "GesCtb_frm_927.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Aceptar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_927.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   8835
            Picture         =   "GesCtb_frm_927.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   765
         Left            =   30
         TabIndex        =   14
         Top             =   1470
         Width           =   9405
         _Version        =   65536
         _ExtentX        =   16589
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
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1530
            MaxLength       =   250
            TabIndex        =   1
            Text            =   "txt_Descri"
            Top             =   390
            Width           =   7665
         End
         Begin VB.ComboBox cmb_buscar 
            Height          =   315
            ItemData        =   "GesCtb_frm_927.frx":1076
            Left            =   1530
            List            =   "GesCtb_frm_927.frx":1078
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3465
         End
         Begin VB.Label Label2 
            Caption         =   "Busca Por:"
            Height          =   225
            Left            =   60
            TabIndex        =   16
            Top             =   90
            Width           =   1515
         End
         Begin VB.Label Label7 
            Caption         =   "Descripción:"
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   435
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_AsiCtb_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_Buscar_Click()
   Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmb_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Call gs_SetFocus(txt_Descri)
End Sub

Private Sub cmd_Acepta_Click()
   If grd_Listad.Rows > 0 Then
      frm_Ctb_AsiCtb_02.txt_CtaCtb.Text = Empty
      grd_Listad.Col = 0
      frm_Ctb_AsiCtb_02.txt_CtaCtb.Text = CStr(grd_Listad.Text)
      Call cmd_Limpia_Click
      frm_Ctb_AsiCtb_04.Hide
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_Buscar.ListIndex = -1 Then
      MsgBox "Debe seleccionar búsqueda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Buscar)
      Exit Sub
   End If
   
   If txt_Descri.Text = Empty Then
      MsgBox "Debe ingresar datos a buscar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If
     
   Screen.MousePointer = 11
   Call fs_Activa(False)
   Call fs_Buscar
   cmb_Buscar.Enabled = False
   txt_Descri.Enabled = False
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Activa(True)
   cmb_Buscar.Enabled = True
   txt_Descri.Enabled = True
   Call gs_SetFocus(cmb_Buscar)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 2175
   grd_Listad.ColWidth(1) = 6795
 
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_Limpia()
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "CUENTA"
   cmb_Buscar.ItemData(cmb_Buscar.NewIndex) = 1
   cmb_Buscar.AddItem "DESCRIPCION"
   cmb_Buscar.ItemData(cmb_Buscar.NewIndex) = 2
   
   cmb_Buscar.ListIndex = 0
   txt_Descri.Text = Empty
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmd_Buscar.Enabled = p_Activa
   cmd_Limpia.Enabled = Not p_Activa
   cmd_Acepta.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   
   Call fs_Activa(False)
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT TRIM(CNTA_CTBL) CNTA_CTBL, TRIM(DESC_CNTA) DESC_CNTA "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CNTA WHERE LENGTH (TRIM(CNTA_CTBL)) = 12 AND"
   If Me.cmb_Buscar.ItemData(cmb_Buscar.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & " CNTA_CTBL LIKE  '" & txt_Descri.Text & "%'"
   Else
      g_str_Parame = g_str_Parame & " DESC_CNTA LIKE '" & txt_Descri.Text & "%'"
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY CNTA_CTBL ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Call gs_SetFocus(cmd_Limpia)
     Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!CNTA_CTBL)
      
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(g_rst_Princi!DESC_CNTA)
           
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
     
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(cmd_Limpia) 'grd_Listad
End Sub
Private Sub grd_Listad_DblClick()
'   If grd_Listad.Rows > 0 Then
'      Call cmd_Acepta_Click
'   End If
End Sub
Private Sub grd_Listad_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
           cmd_Acepta_Click
   End Select
End Sub
Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub




