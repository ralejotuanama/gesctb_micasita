VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Mnt_EFBG_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "GesCtb_frm_909.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7485
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7485
      _Version        =   65536
      _ExtentX        =   13203
      _ExtentY        =   13203
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
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de EEFF - Balance General"
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
            Picture         =   "GesCtb_frm_909.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3405
         Left            =   60
         TabIndex        =   12
         Top             =   2355
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   5997
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
            Width           =   5595
            _Version        =   65536
            _ExtentX        =   9869
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
            _ExtentX        =   2566
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
         Left            =   60
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_909.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Cancelar "
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_909.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6750
            Picture         =   "GesCtb_frm_909.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_909.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_909.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_909.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   825
         Left            =   60
         TabIndex        =   16
         Top             =   1470
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
         Begin Threed.SSPanel pnl_Grp 
            Height          =   315
            Left            =   1800
            TabIndex        =   17
            Top             =   90
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_SubGrp 
            Height          =   315
            Left            =   1800
            TabIndex        =   18
            Top             =   420
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   556
            _StockProps     =   15
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
            Caption         =   "Grupo :"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   90
            Width           =   1635
         End
         Begin VB.Label Label1 
            Caption         =   "Sub Grupo :"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   450
            Width           =   1635
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   1605
         Left            =   60
         TabIndex        =   21
         Top             =   5820
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
         _ExtentY        =   2831
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
         Begin VB.TextBox txt_Agrupa 
            Height          =   315
            Left            =   1860
            MaxLength       =   5
            TabIndex        =   26
            Top             =   1080
            Width           =   660
         End
         Begin VB.TextBox txt_TipCta 
            Height          =   315
            Left            =   1860
            MaxLength       =   1
            TabIndex        =   24
            Top             =   750
            Width           =   660
         End
         Begin VB.TextBox txt_Descri 
            Height          =   315
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   8
            Top             =   420
            Width           =   5445
         End
         Begin VB.TextBox txt_Codigo 
            Height          =   315
            Left            =   1860
            MaxLength       =   12
            TabIndex        =   7
            Top             =   90
            Width           =   1875
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Agrupación:"
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   1110
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Tipo de Cuenta:"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   780
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Descripción:"
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   450
            Width           =   1485
         End
         Begin VB.Label lbl_Etique 
            Caption         =   "Código :"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_EFBG_02"
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
   g_str_Parame = g_str_Parame & "SELECT TO_CHAR(CNTA_CTBLE) CNTA_CTBLE, NOMB_CTACTB, TIPO_CTACTB, TIPO_AGRP "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_EEBG "
   g_str_Parame = g_str_Parame & " WHERE TRIM(INDC_TIPO) = 'D' "
   g_str_Parame = g_str_Parame & "   AND CODG_GRUPO = " & moddat_g_str_Codigo & "  "
   g_str_Parame = g_str_Parame & "   AND CODG_SBGRP = " & CInt(moddat_g_str_CodIte) & "  "
   g_str_Parame = g_str_Parame & "   AND CNTA_CTBLE = " & r_int_NroCta & "  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
    
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
    
   g_rst_Princi.MoveFirst
   txt_Codigo.Text = g_rst_Princi!CNTA_CTBLE
   txt_Descri.Text = Trim(g_rst_Princi!NOMB_CTACTB)
   txt_TipCta.Text = Trim(g_rst_Princi!TIPO_CTACTB)
   If IsNull(g_rst_Princi!TIPO_AGRP) Then
      txt_Agrupa.Text = ""
   Else
      txt_Agrupa.Text = Trim(g_rst_Princi!TIPO_AGRP)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   txt_Codigo.Enabled = False
   Call gs_SetFocus(txt_Descri)
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
   g_str_Parame = "USP_CTB_EEFF_BG ("
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'" & r_str_CodSubGrp & "', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & 3 & ","
   g_str_Parame = g_str_Parame & 2 & ")"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
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
   If Len(Trim(txt_TipCta.Text)) = 0 Then
      MsgBox "Debe ingresar el tipo de cuenta contable.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_TipCta)
      Exit Sub
   End If
   If Not (Trim(txt_TipCta.Text) <> "D" Or Trim(txt_TipCta.Text) <> "H") Then
      MsgBox "El tipo de cuenta solo puede ser 'D'-Debe o 'H'-Haber.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_TipCta)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   Screen.MousePointer = 11
   
   If moddat_g_int_FlgGrb = 1 Then
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = ""
         g_str_Parame = "USP_CTB_EEFF_BG ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(r_str_DescGrp) & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(r_str_DescSubGrp) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Codigo.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_TipCta.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Descri.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Agrupa.Text) & "', "
         g_str_Parame = g_str_Parame & 1 & ","
         g_str_Parame = g_str_Parame & 2 & ")"
               
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Screen.MousePointer = 0
               Exit Sub
            Else
               moddat_g_int_FlgGOK = True
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
   Else
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CTB_EEFF_BG ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_Codigo & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Codigo.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_TipCta.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Descri.Text) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(txt_Agrupa.Text) & "', "
         g_str_Parame = g_str_Parame & 2 & ","
         g_str_Parame = g_str_Parame & 2 & ")"
            
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
   
         If moddat_g_int_CntErr > 0 Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               Screen.MousePointer = 0
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
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   frm_Mnt_EFBG_01.fs_Buscar
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
   
   Call gs_SetFocus(cmd_Agrega)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Buscar()
   Call gs_LimpiaGrid(grd_Listad)
     
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT CODG_GRUPO, NOMB_GRUPO, CODG_SBGRP, NOMB_SBGRP, "
   g_str_Parame = g_str_Parame & "       TO_CHAR(CNTA_CTBLE) CNTA_CTBLE, NOMB_CTACTB, INDC_TIPO, ITEM_CNTA, TIPO_AGRP "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_EEBG "
   g_str_Parame = g_str_Parame & " WHERE TRIM(INDC_TIPO) = 'D' "
   g_str_Parame = g_str_Parame & "   AND CODG_GRUPO = " & moddat_g_str_Codigo & "  "
   g_str_Parame = g_str_Parame & "   AND CODG_SBGRP = " & CInt(moddat_g_str_CodIte) & "  "
   g_str_Parame = g_str_Parame & "   AND CNTA_CTBLE IS NOT NULL "
   g_str_Parame = g_str_Parame & " ORDER BY ITEM_CNTA "
    
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
   
   pnl_Grp.Caption = moddat_g_str_Codigo & " - " & moddat_g_str_Descri
   pnl_SubGrp.Caption = moddat_g_str_CodIte & " - " & moddat_g_str_DesIte
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
   txt_Agrupa.Enabled = p_Activa
   SSPanel6.Enabled = p_Activa
End Sub

Private Sub fs_Limpia()
   txt_Codigo.Text = ""
   txt_Descri.Text = ""
   txt_TipCta.Text = ""
   txt_Agrupa.Text = ""
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

Private Sub txt_Codigo_GotFocus()
   Call gs_SelecTodo(txt_Codigo)
End Sub

Private Sub txt_Codigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descri)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Descri_GotFocus()
   Call gs_SelecTodo(txt_Descri)
End Sub

Private Sub txt_Descri_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_TipCta)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & " ,-/\[]().")
   End If
End Sub

Private Sub txt_TipCta_GotFocus()
   Call gs_SelecTodo(txt_TipCta)
End Sub

Private Sub txt_TipCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Agrupa)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, "HD")
   End If
End Sub

Private Sub txt_Agrupa_GotFocus()
   Call gs_SelecTodo(txt_Agrupa)
End Sub

Private Sub txt_Agrupa_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, "0123456789")
   End If
End Sub

