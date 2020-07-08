VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_Person_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   7110
   ClientLeft      =   4635
   ClientTop       =   2460
   ClientWidth     =   10260
   Icon            =   "GesCtb_frm_830.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10275
      _Version        =   65536
      _ExtentX        =   18124
      _ExtentY        =   12568
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
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   10185
         _Version        =   65536
         _ExtentX        =   17965
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   525
            Left            =   600
            TabIndex        =   2
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "Mantenimiento de Personal"
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
            Picture         =   "GesCtb_frm_830.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   750
         Width           =   10185
         _Version        =   65536
         _ExtentX        =   17965
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_830.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Borrar Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_830.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Nueva Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_830.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Modificar Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9570
            Picture         =   "GesCtb_frm_830.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Ventana"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5595
         Left            =   30
         TabIndex        =   8
         Top             =   1470
         Width           =   10185
         _Version        =   65536
         _ExtentX        =   17965
         _ExtentY        =   9869
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
            Height          =   5175
            Left            =   30
            TabIndex        =   9
            Top             =   360
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   9128
            _Version        =   393216
            Rows            =   12
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_CodEmp 
            Height          =   285
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   2085
            _Version        =   65536
            _ExtentX        =   3678
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
         Begin Threed.SSPanel pnl_Tit_RazSoc 
            Height          =   285
            Left            =   2130
            TabIndex        =   11
            Top             =   60
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Completo"
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
            Left            =   7980
            TabIndex        =   12
            Top             =   60
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Categoría"
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
Attribute VB_Name = "frm_Mnt_Person_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim r_int_PerMes        As String
   Dim r_int_PerAno        As String

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   frm_Mnt_Person_02.Show 1
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
     
   'Instrucción SQL
   g_str_Parame = "DELETE FROM CTB_USUSBS WHERE USUSBS_PERMES = " & r_int_PerMes & " AND USUSBS_PERANO = " & r_int_PerAno & " AND "
   g_str_Parame = g_str_Parame & "USUSBS_CODUSU = '" & CStr(Trim(grd_Listad.Text)) & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   'MsgBox "Registro eliminado.", vbInformation, modgen_g_str_NomPlt
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_int_FlgGrb = 2
   grd_Listad.Col = 0
   modsec_g_str_CodUsu = CStr(Trim(grd_Listad.Text))
   frm_Mnt_Person_02.Show 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Call ff_BusPer
   Call fs_Buscar
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt

   Call fs_Inicio
   
   Call gs_CentraForm(Me)
      
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   grd_Listad.ColWidth(0) = 2075
   grd_Listad.ColWidth(1) = 5855
   grd_Listad.ColWidth(2) = 1775
   
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
   
   g_str_Parame = "SELECT * FROM CTB_USUSBS WHERE USUSBS_PERMES = " & r_int_PerMes & " AND USUSBS_PERANO = " & r_int_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY "
   g_str_Parame = g_str_Parame & "USUSBS_APEPAT, USUSBS_APEMAT ASC, USUSBS_NOMBRE ASC"
   
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
      grd_Listad.Text = Trim(g_rst_Princi!USUSBS_CODUSU)
      
      grd_Listad.Col = 1
      
      grd_Listad.Text = Trim(g_rst_Princi!USUSBS_APEPAT) & " " & Trim(g_rst_Princi!USUSBS_APEMAT) & " " & Trim(g_rst_Princi!USUSBS_NOMBRE)
      
      grd_Listad.Col = 2
      
      If CStr(g_rst_Princi!USUSBS_TIPPER) = "1" Then
         grd_Listad.Text = "GERENTE"
      ElseIf CStr(g_rst_Princi!USUSBS_TIPPER) = "2" Then
         grd_Listad.Text = "FUNCIONARIO"
      ElseIf CStr(g_rst_Princi!USUSBS_TIPPER) = "3" Then
         grd_Listad.Text = "EMPLEADO"
      ElseIf CStr(g_rst_Princi!USUSBS_TIPPER) = "4" Then
         grd_Listad.Text = "OTROS"
         
      End If
      
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

Public Sub ff_BusPer()
      
   g_str_Parame = ""
   g_str_Parame = "SELECT MAX(USUSBS_PERMES) AS PERMES, MAX(USUSBS_PERANO) AS PERANO FROM CTB_USUSBS "
   
   If Format(Now, "MM") - 1 = 0 Then
      g_str_Parame = g_str_Parame & "WHERE USUSBS_PERANO = " & Format(Now, "YYYY") - 1
   Else
      g_str_Parame = g_str_Parame & "WHERE USUSBS_PERANO = " & Format(Now, "YYYY")
   End If
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
       r_int_PerMes = g_rst_Listas!PERMES
       r_int_PerAno = g_rst_Listas!PERANO
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub


