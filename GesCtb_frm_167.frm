VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mat_ConCtb_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7605
   ClientLeft      =   4905
   ClientTop       =   2865
   ClientWidth     =   10590
   Icon            =   "GesCtb_frm_167.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7575
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10575
      _Version        =   65536
      _ExtentX        =   18653
      _ExtentY        =   13361
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
         TabIndex        =   6
         Top             =   60
         Width           =   10485
         _Version        =   65536
         _ExtentX        =   18494
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
            TabIndex        =   7
            Top             =   60
            Width           =   6375
            _Version        =   65536
            _ExtentX        =   11245
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Matriz Contable - Conceptos Contables"
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
            Picture         =   "GesCtb_frm_167.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   6045
         Left            =   30
         TabIndex        =   8
         Top             =   1470
         Width           =   10485
         _Version        =   65536
         _ExtentX        =   18494
         _ExtentY        =   10663
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
            Height          =   5655
            Left            =   30
            TabIndex        =   0
            Top             =   360
            Width           =   10425
            _ExtentX        =   18389
            _ExtentY        =   9975
            _Version        =   393216
            Rows            =   25
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Descrip 
            Height          =   285
            Left            =   1680
            TabIndex        =   9
            Top             =   60
            Width           =   5775
            _Version        =   65536
            _ExtentX        =   10186
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Descripci�n"
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
         Begin Threed.SSPanel pnl_Campo 
            Height          =   285
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "C�digo Campo"
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
         Begin Threed.SSPanel pnl_TipCon 
            Height          =   285
            Left            =   7440
            TabIndex        =   12
            Top             =   60
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
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
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   11
         Top             =   780
         Width           =   10485
         _Version        =   65536
         _ExtentX        =   18494
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9870
            Picture         =   "GesCtb_frm_167.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_167.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_167.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_167.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Mat_ConCtb_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_FlgAct = 1
   
   frm_Mat_ConCtb_02.Show 1
   
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

   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)

   If MsgBox("�Est� seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Instrucci�n SQL
   g_str_Parame = "DELETE FROM CTB_CONCTB WHERE "
   g_str_Parame = g_str_Parame & "CONCTB_CODCAM = '" & moddat_g_str_CodGrp & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   moddat_g_str_CodGrp = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   moddat_g_int_FlgAct = 1
   
   frm_Mat_ConCtb_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt

   Call fs_Inicio
   Call gs_CentraForm(Me)
   Call fs_Buscar

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   grd_Listad.ColWidth(0) = 1625
   grd_Listad.ColWidth(1) = 5765
   grd_Listad.ColWidth(2) = 2655
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar()
   cmd_Agrega.Enabled = True
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
   grd_Listad.Enabled = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = "SELECT * FROM CTB_CONCTB ORDER BY CONCTB_CODCAM ASC"
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
      grd_Listad.Text = Trim(g_rst_Princi!CONCTB_CODCAM)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!CONCTB_DESCRI)
      
      grd_Listad.Col = 2
      grd_Listad.Text = moddat_gf_Consulta_ParDes("063", CStr(g_rst_Princi!CONCTB_TIPCON))
      
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

Private Sub pnl_Campo_Click()
   If pnl_Campo.Tag = "" Then
      pnl_Campo.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 0, "N")
   Else
      pnl_Campo.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 0, "N-")
   End If
End Sub

Private Sub pnl_Descrip_Click()
   If pnl_Descrip.Tag = "" Then
      pnl_Descrip.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Descrip.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_TipCon_Click()
   If pnl_TipCon.Tag = "" Then
      pnl_TipCon.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_TipCon.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub
