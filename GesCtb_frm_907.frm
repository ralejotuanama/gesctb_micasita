VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_EFGP_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   Icon            =   "GesCtb_frm_907.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6680
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
      _ExtentY        =   11783
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.26
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   10
         Top             =   60
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
            TabIndex        =   11
            Top             =   60
            Width           =   6675
            _Version        =   65536
            _ExtentX        =   11774
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de EEFF - Estados de Ganancias y Pérdidas"
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
            Picture         =   "GesCtb_frm_907.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3400
         Left            =   30
         TabIndex        =   12
         Top             =   2325
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   5997
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   3000
            Left            =   30
            TabIndex        =   6
            Top             =   360
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   5292
            _Version        =   393216
            Rows            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_Descri 
            Height          =   285
            Left            =   1440
            TabIndex        =   13
            Top             =   60
            Width           =   5580
            _Version        =   65536
            _ExtentX        =   9842
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
         Begin Threed.SSPanel pnl_Tit_Codigo 
            Height          =   285
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2575
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
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   15
         Top             =   780
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_907.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_907.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_907.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6750
            Picture         =   "GesCtb_frm_907.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   1800
            Picture         =   "GesCtb_frm_907.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   2400
            Picture         =   "GesCtb_frm_907.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Cancelar "
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   795
         Left            =   30
         TabIndex        =   16
         Top             =   1470
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   1402
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
         Begin Threed.SSPanel pnl_Grp 
            Height          =   315
            Left            =   1800
            TabIndex        =   21
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_SubGrp 
            Height          =   315
            Left            =   1800
            TabIndex        =   23
            Top             =   420
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin VB.Label Label1 
            Caption         =   "Sub Grupo :"
            Height          =   255
            Left            =   90
            TabIndex        =   22
            Top             =   450
            Width           =   1665
         End
         Begin VB.Label Label4 
            Caption         =   "Grupo :"
            Height          =   255
            Left            =   90
            TabIndex        =   17
            Top             =   90
            Width           =   1665
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   795
         Left            =   30
         TabIndex        =   18
         Top             =   5790
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   1402
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
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1860
            MaxLength       =   30
            TabIndex        =   7
            Top             =   60
            Width           =   1875
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   8
            Top             =   390
            Width           =   5445
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Código :"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   20
            Top             =   90
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Descripción:"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   19
            Top             =   420
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_EFGP_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_int_NroCta        As String
Dim r_str_CodSubGrp     As String
Dim r_str_DescGrp       As String
Dim r_str_DescSubGrp    As String

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   
   Screen.MousePointer = 11
   Call fs_Activa(True)
   Call fs_Limpia
   Call gs_SetFocus(txt_Codigo)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Borrar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   r_str_CodSubGrp = grd_Listad.Text
   Call gs_RefrescaGrid(grd_Listad)
    
   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
    
   'Instrucción SQL
   g_str_Parame = ""
   g_str_Parame = "USP_CTB_EEFF_EGP ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & r_str_CodSubGrp & ", "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & 3 & ","
   g_str_Parame = g_str_Parame & 2 & ")"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
    
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Activa(False)
   Call fs_Limpia
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_int_FlgGrb = 2
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
     
   r_int_NroCta = CStr(grd_Listad.Text)
   Screen.MousePointer = 11
   Call fs_Activa(True)
   Call fs_Limpia
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT to_char(CNTA_CTBLE) CNTA_CTBLE,NOMB_CTACTB "
   g_str_Parame = g_str_Parame & "From CNTBL_EEFF "
   g_str_Parame = g_str_Parame & "WHERE TRIM(INDC_TIPO) = 'D' "
   g_str_Parame = g_str_Parame & "and CODG_GRUPO=" & moddat_g_str_Codigo & "  "
   g_str_Parame = g_str_Parame & "and CODG_SBGRP=" & CInt(moddat_g_str_CodIte) & "  "
   g_str_Parame = g_str_Parame & "and CNTA_CTBLE=" & r_int_NroCta & "  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
    
   g_rst_Princi.MoveFirst
   txt_Codigo.Text = g_rst_Princi!CNTA_CTBLE
   txt_Descri.Text = Trim(g_rst_Princi!NOMB_CTACTB)
   
   txt_Codigo.Enabled = False
   Call gs_SetFocus(txt_Descri)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_Codigo.Text)) = 0 Then
      MsgBox "Debe ingresar el Código de Clasificación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Codigo)
      Exit Sub
   End If
   If Len(Trim(txt_Descri.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descri)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   If moddat_g_int_FlgGrb = 1 Then
   
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = ""
         g_str_Parame = "USP_CTB_EEFF_EGP ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(r_str_DescGrp) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(r_str_DescSubGrp) & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
         g_str_Parame = g_str_Parame & 1 & ","
         g_str_Parame = g_str_Parame & 2 & ")"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
        
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_FlgGOK = True
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   
   Else
   
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CTB_EEFF_EGP ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & txt_Codigo.Text & "', "
         g_str_Parame = g_str_Parame & "'" & txt_Descri.Text & "', "
         g_str_Parame = g_str_Parame & 2 & ","
         g_str_Parame = g_str_Parame & 2 & ")"
            
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_FlgGOK = True
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   End If
   
   Call fs_Buscar
   Call fs_Activa(False)
   Call fs_Limpia
End Sub

Private Sub cmd_Salida_Click()
   frm_Mnt_EFGP_01.fs_Buscar
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   r_str_DescGrp = moddat_g_str_Descri
   
   Call fs_Inicia
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Activa(False)
   Call fs_Limpia
   Call fs_Buscar
      
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Buscar()
   pnl_Grp.Caption = ""
   pnl_SubGrp.Caption = ""
   Call gs_LimpiaGrid(grd_Listad)
     
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT CODG_GRUPO, NOMB_GRUPO, CODG_SBGRP, NOMB_SBGRP,"
   g_str_Parame = g_str_Parame & "       TO_CHAR(CNTA_CTBLE) CNTA_CTBLE, NOMB_CTACTB, INDC_TIPO, ITEM_CNTA "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_EEFF "
   g_str_Parame = g_str_Parame & " WHERE TRIM(INDC_TIPO) = 'D' "
   g_str_Parame = g_str_Parame & "   AND CODG_GRUPO       = " & moddat_g_str_Codigo & "  "
   g_str_Parame = g_str_Parame & "   AND CODG_SBGRP       = " & CInt(moddat_g_str_CodIte) & "  "
   g_str_Parame = g_str_Parame & "   AND TRIM(NOMB_GRUPO) = '" & moddat_g_str_Descri & "' "
   g_str_Parame = g_str_Parame & "   AND CNTA_CTBLE IS NOT NULL "
   g_str_Parame = g_str_Parame & "ORDER BY ITEM_CNTA"
    
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
    
   pnl_Grp.Caption = g_rst_Princi!CODG_GRUPO & " - " & g_rst_Princi!NOMB_GRUPO
   pnl_SubGrp.Caption = g_rst_Princi!CODG_SBGRP & " - " & g_rst_Princi!NOMB_SBGRP
   r_str_DescGrp = g_rst_Princi!NOMB_GRUPO
   r_str_DescSubGrp = g_rst_Princi!NOMB_SBGRP
    
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = g_rst_Princi!CNTA_CTBLE
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!NOMB_CTACTB)
      
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

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1365
   grd_Listad.ColWidth(1) = 5595
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   moddat_g_int_FlgGrb = 0
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   grd_Listad.Enabled = Not p_Activa
   cmd_Agrega.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   txt_Codigo.Enabled = p_Activa
   txt_Descri.Enabled = p_Activa
   SSPanel6.Enabled = p_Activa
End Sub

Private Sub fs_Limpia()
   txt_Codigo.Text = ""
   txt_Descri.Text = ""
End Sub

Private Sub pnl_Tit_Codigo_Click()
   If Len(Trim(pnl_Tit_Codigo.Tag)) = 0 Or pnl_Tit_Codigo.Tag = "D" Then
      pnl_Tit_Codigo.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_Codigo.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_Descri_Click()
   If Len(Trim(pnl_Tit_Descri.Tag)) = 0 Or pnl_Tit_Descri.Tag = "D" Then
      pnl_Tit_Descri.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_Descri.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descri)
   End If
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub


