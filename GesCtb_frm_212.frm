VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_CarArc_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   Icon            =   "GesCtb_frm_212.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6510
      Left            =   -30
      TabIndex        =   14
      Top             =   0
      Width           =   10995
      _Version        =   65536
      _ExtentX        =   19394
      _ExtentY        =   11483
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
         Left            =   90
         TabIndex        =   15
         Top             =   60
         Width           =   10820
         _Version        =   65536
         _ExtentX        =   19085
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
         Begin Threed.SSPanel pnl_Titulo 
            Height          =   300
            Left            =   660
            TabIndex        =   16
            Top             =   150
            Width           =   6225
            _Version        =   65536
            _ExtentX        =   10980
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Carga del Archivo de Recaudo"
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
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   10320
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_212.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   90
         TabIndex        =   17
         Top             =   780
         Width           =   10820
         _Version        =   65536
         _ExtentX        =   19085
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   30
            Picture         =   "GesCtb_frm_212.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10215
            Picture         =   "GesCtb_frm_212.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2535
         Left            =   90
         TabIndex        =   18
         Top             =   3880
         Width           =   10820
         _Version        =   65536
         _ExtentX        =   19085
         _ExtentY        =   4471
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
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1380
            Width           =   2880
         End
         Begin VB.ComboBox cmb_CtaCte 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2055
            Width           =   2880
         End
         Begin VB.ComboBox cmb_Banco 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1710
            Width           =   2880
         End
         Begin VB.TextBox txt_Descrip 
            Height          =   315
            Left            =   1530
            MaxLength       =   60
            TabIndex        =   8
            Top             =   1050
            Width           =   6390
         End
         Begin VB.ComboBox cmb_Proveedor 
            Height          =   315
            Left            =   1530
            TabIndex        =   7
            Top             =   720
            Width           =   6390
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   390
            Width           =   6390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   780
            Width           =   780
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   450
            Width           =   1230
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   90
            Width           =   885
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1110
            Width           =   885
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1770
            Width           =   510
         End
         Begin VB.Label lbl_Cuenta 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Corriente:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   2100
            Width           =   1230
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   825
         Left            =   90
         TabIndex        =   25
         Top             =   1470
         Width           =   10820
         _Version        =   65536
         _ExtentX        =   19085
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   315
            Left            =   1530
            TabIndex        =   33
            Top             =   390
            Width           =   1600
            _Version        =   65536
            _ExtentX        =   2822
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_FecProc 
            Height          =   315
            Left            =   4980
            TabIndex        =   35
            Top             =   390
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   90
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Proceso:"
            Height          =   195
            Left            =   3780
            TabIndex        =   36
            Top             =   450
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   450
            Width           =   540
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   1500
         Left            =   90
         TabIndex        =   26
         Top             =   2340
         Width           =   10820
         _Version        =   65536
         _ExtentX        =   19085
         _ExtentY        =   2646
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
         Begin VB.TextBox txt_NomArc 
            Height          =   315
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   405
            Width           =   7935
         End
         Begin VB.CommandButton cmd_BuscaArc 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9480
            TabIndex        =   0
            ToolTipText     =   "Seleccionar archivo"
            Top             =   390
            Width           =   315
         End
         Begin VB.CommandButton cmd_Import 
            Height          =   585
            Left            =   10215
            Picture         =   "GesCtb_frm_212.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Importar archivo"
            Top             =   120
            Width           =   585
         End
         Begin Threed.SSPanel pnl_FecAch 
            Height          =   315
            Left            =   1530
            TabIndex        =   2
            Top             =   750
            Width           =   1600
            _Version        =   65536
            _ExtentX        =   2822
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_MonAch 
            Height          =   315
            Left            =   4980
            TabIndex        =   3
            Top             =   750
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_FilAch 
            Height          =   315
            Left            =   1530
            TabIndex        =   5
            Top             =   1090
            Width           =   1600
            _Version        =   65536
            _ExtentX        =   2822
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_TotAch 
            Height          =   315
            Left            =   7860
            TabIndex        =   4
            Top             =   750
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Archivo a cargar:"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   450
            Width           =   1215
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   3780
            TabIndex        =   32
            Top             =   810
            Width           =   630
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   810
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Resumen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   90
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Importe Total:"
            Height          =   195
            Left            =   6810
            TabIndex        =   28
            Top             =   810
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nro Filas:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1140
            Width           =   660
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_CarArc_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type arr_RegCli
   r_str_TipDoc    As String
   r_str_NumDoc    As String
   r_str_NumOpe    As String
   r_dbl_Import    As Double
   r_str_NomCli    As String
   r_int_NroCta    As Integer
   r_str_TipPag    As String
   r_str_CodMon    As Integer
End Type
   
Dim l_arr_GenArc()      As arr_RegCli
Dim l_arr_MaePrv()      As moddat_tpo_Genera
Dim l_arr_CtaCteSol()   As moddat_tpo_Genera
Dim l_arr_CtaCteDol()   As moddat_tpo_Genera
Dim l_int_Contar        As Integer
Dim l_int_TipBnc        As Integer

Private Sub cmb_Moneda_Click()
   Call fs_CargarBancos
End Sub

Private Sub cmb_Proveedor_Click()
   Call fs_Buscar_prov
End Sub

Private Sub cmb_TipDoc_Click()
   Call fs_CargarPrv
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt

   Call fs_Inicia
   Call fs_Limpiar
   Call gs_CentraForm(Me)
   
   cmd_Grabar.Visible = False
   If moddat_g_int_FlgGrb = 0 Then
      pnl_Titulo.Caption = "Carga del Archivo de Recaudo - Consulta"
      Call fs_Cargar_Datos
      Call fs_Desabilitar
   ElseIf moddat_g_int_FlgGrb = 1 Then
      pnl_Titulo.Caption = "Carga del Archivo de Recaudo - Adicionar"
      pnl_FecProc.Caption = Format(moddat_g_str_FecSis, "dd/mm/yyyy")
      'pnl_FecProc.Caption = "31/12/2018"
      cmd_Grabar.Visible = True
      cmd_Grabar.Enabled = False
      cmd_Import.Enabled = False
   End If
   
   Call gs_SetFocus(txt_NomArc)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_BuscaArc_Click()
Dim r_bol_Estado As Boolean

   r_bol_Estado = False
   
   dlg_Guarda.FileName = ""
   dlg_Guarda.Filter = "Archivo de Texto (*.txt)|*.txt"
   dlg_Guarda.ShowOpen
   If Trim(dlg_Guarda.FileName) <> "" Then
      txt_NomArc.Text = UCase(dlg_Guarda.FileName)
      cmd_Grabar.Enabled = False
      pnl_FecAch.Caption = ""
      pnl_MonAch.Caption = ""
      pnl_TotAch.Caption = "0.00" & " "
      pnl_FilAch.Caption = "0"
      ReDim l_arr_GenArc(0)
   End If

   If UCase(Mid(dlg_Guarda.FileTitle, 9, 17)) = "_REC00420PGNB.TXT" Or UCase(Mid(dlg_Guarda.FileTitle, 9, 17)) = "_REC00421PGNB.TXT" Or _
      UCase(Mid(dlg_Guarda.FileTitle, 9, 17)) = "_REC00420PCOM.TXT" Or UCase(Mid(dlg_Guarda.FileTitle, 9, 17)) = "_REC00421PCOM.TXT" Then
      r_bol_Estado = True
      cmd_Import.Enabled = True
      
      l_int_TipBnc = 0
      If UCase(Mid(dlg_Guarda.FileTitle, 18, 4)) = "PGNB" Then
         l_int_TipBnc = 1 '"BANCO GNB"
      Else
         l_int_TipBnc = 2 '"BANCO COMERCIO"
      End If
   End If
   
   If r_bol_Estado = False Then
      l_int_TipBnc = 0
      pnl_FecAch.Caption = ""
      pnl_MonAch.Caption = ""
      pnl_TotAch.Caption = "0.00" & " "
      pnl_FilAch.Caption = "0"
      ReDim l_arr_GenArc(0)
      cmd_Import.Enabled = False
      txt_NomArc.Text = ""
      MsgBox "El archivo seleccionado no cumple con el formato.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If
   
   Call gs_SetFocus(cmd_Import)
   Exit Sub
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Limpiar()
   l_int_TipBnc = 0
   pnl_Codigo.Caption = ""
   pnl_FilAch.Caption = "0"
   pnl_TotAch.Caption = "0.00" & " "
   pnl_FecAch.Caption = ""
   pnl_MonAch.Caption = ""
   cmb_TipDoc.ListIndex = -1
   cmb_Proveedor.ListIndex = -1
   txt_Descrip.Text = ""
   cmb_Banco.ListIndex = -1
   cmb_Moneda.ListIndex = -1
   cmb_CtaCte.ListIndex = -1
   Call gs_SetFocus(cmb_TipDoc)
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   'Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "118")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
End Sub

Private Sub fs_Desabilitar()
   txt_NomArc.Enabled = False
   cmd_BuscaArc.Enabled = False
   cmd_Import.Enabled = False
   cmb_TipDoc.Enabled = False
   cmb_Proveedor.Enabled = False
   txt_Descrip.Enabled = False
   cmb_Moneda.Enabled = False
   cmb_Banco.Enabled = False
   cmb_CtaCte.Enabled = False
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_Msg    As String
Dim r_bol_Estado As Boolean
Dim r_dbl_TipSbs As Boolean

   If Trim(txt_NomArc.Text) = "" Then
      MsgBox "Debe ingresar la ubicación y nombre del archivo, luego darle click en boton importar .", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If
      
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un tipo de documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(cmb_Proveedor.Text)) = 0 Then
       MsgBox "Tiene que ingresar un proveedor.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Proveedor)
       Exit Sub
   Else
       If (fs_ValNumDoc() = False) Then
           Exit Sub
       Else
           r_bol_Estado = False
           If InStr(1, Trim(cmb_Proveedor.Text), "-") > 0 Then
              For l_int_Contar = 1 To UBound(l_arr_MaePrv)
                  If Trim(Mid(cmb_Proveedor.Text, 1, InStr(Trim(cmb_Proveedor.Text), "-") - 1)) = Trim(l_arr_MaePrv(l_int_Contar).Genera_Codigo) Then
                     r_bol_Estado = True
                     Exit For
                  End If
              Next
           End If
           If r_bol_Estado = False Then
              MsgBox "El Proveedor no se encuentra en la lista.", vbExclamation, modgen_g_str_NomPlt
              Call gs_SetFocus(cmb_Proveedor)
              Exit Sub
           End If
       End If
   End If
   '------------------
   If cmb_Moneda.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Moneda)
      Exit Sub
   End If
      
   If Len(Trim(txt_Descrip.Text)) = 0 Then
      MsgBox "Tiene que ingresar una descripción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Descrip)
      Exit Sub
   End If
   
   If cmb_Banco.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Banco)
      Exit Sub
   End If
   
   If cmb_CtaCte.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un nro cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaCte)
      Exit Sub
   End If
   
   If CDbl(pnl_TotAch.Caption) <= 0 Then
      MsgBox "El importe a pagar no puede ser cero ni negativo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If
         
   If Format(pnl_FecProc.Caption, "yyyymm") <> modctb_int_PerAno & Format(modctb_int_PerMes, "00") Then
      MsgBox "El documento se encuentra fuera del periodo actual.", vbExclamation, modgen_g_str_NomPlt
            
      If MsgBox("¿Esta seguro de registrar un documento fuera del periodo actual?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Call gs_SetFocus(pnl_FecProc)
         Exit Sub
      End If
   End If
   
   'TipCam = 1 - Comercial / 2 - SBS / 3 - Sunat / 4 - BCR
   'TipTip = 1 - Venta / 2 - Compra
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, Format(Trim(pnl_FecProc.Caption), "yyyymmdd"), 1)
   If r_dbl_TipSbs = 0 Then
      MsgBox "Falta definir el tipo de cambio SBS del día " & Format(date, "dd/mm/yyyy") & ".", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Grabar)
      Exit Sub
   End If
   
   r_str_Msg = ""
   If l_int_TipBnc = 0 Then
      MsgBox "En el nombre del archivo no esta definido el tipo de banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If
   
'   If (Format(Trim(pnl_FecProc.Caption), "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'       Format(Trim(pnl_FecProc.Caption), "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'       MsgBox "Intenta registrar una operación en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(txt_NomArc)
'       Exit Sub
'   End If
   
   If fs_ValidaPeriodo(pnl_FecProc.Caption) = False Then
      Exit Sub
   End If
   
   r_str_Msg = ""
   If l_int_TipBnc = 1 Then
      r_str_Msg = "Banbo GNB"
   Else
      r_str_Msg = "Banco del Comercio"
   End If
   
   If MsgBox("¿Esta seguro de grabar los datos del " & r_str_Msg & " ?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Grabar
   Screen.MousePointer = 0
   
End Sub

Private Sub fs_Grabar()
Dim r_str_CodGen   As String
Dim r_str_FecPro   As String
Dim r_int_NumPro   As Integer
Dim r_str_AsiGen   As String
   
   Call fs_Guardar_Log(Trim(txt_NomArc.Text), r_str_FecPro, r_int_NumPro)
   
   If Trim(r_str_FecPro) = "" Or r_int_NumPro = 0 Then
      Exit Sub
   End If
   
   r_str_CodGen = ""
   If moddat_g_int_FlgGrb = 1 Then
      r_str_CodGen = modmip_gf_Genera_CodGen(3, 8)
   Else
      r_str_CodGen = Trim(pnl_Codigo.Caption)
   End If

   If Len(Trim(r_str_CodGen)) = 0 Then
      MsgBox "No se genero el código automatico del folio.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_PROPAG_ARH ( "
   g_str_Parame = g_str_Parame & CLng(r_str_CodGen) & ", "
   g_str_Parame = g_str_Parame & Format(Trim(pnl_FecProc.Caption), "yyyymmdd") & ", "
   g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", "
   g_str_Parame = g_str_Parame & "'" & fs_NumDoc(cmb_Proveedor.Text) & "', "
   g_str_Parame = g_str_Parame & "'" & Trim(txt_Descrip.Text) & "', "
   g_str_Parame = g_str_Parame & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", "
   g_str_Parame = g_str_Parame & cmb_Banco.ItemData(cmb_Banco.ListIndex) & ", "
   g_str_Parame = g_str_Parame & "'" & Trim(cmb_CtaCte.Text) & "', "
   g_str_Parame = g_str_Parame & CLng(r_str_FecPro) & ", "
   g_str_Parame = g_str_Parame & CLng(r_int_NumPro) & ", "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   r_str_AsiGen = ""
   If moddat_g_int_FlgGrb = 1 Then
      Call fs_GeneraAsiento(r_str_CodGen, r_str_AsiGen)
      MsgBox "Los datos se grabaron correctamente." & vbCrLf & _
             "El asiento generado es: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
      Call frm_Ctb_CarArc_01.fs_Buscar
      Screen.MousePointer = 0
      Unload Me
   End If
   
End Sub

Private Sub cmd_Import_Click()
Dim r_int_linea          As Integer
Dim r_str_Cadena         As String
Dim r_dbl_Import         As Double
Dim r_str_TipDoc         As String
Dim r_str_NumDoc         As String
Dim r_int_FilAch         As Integer
Dim r_dbl_TotAch         As Double
Dim r_str_FecAch         As String
Dim r_str_MonAch         As String
Dim r_str_CodMon         As Integer

Dim r_int_ValFil         As Integer
Dim r_dbl_ValTot         As Double
Dim r_bol_ValLin         As Boolean
Dim r_bol_ValDNI         As Boolean
Dim r_str_ErrDNI         As String

   If Len(Trim(txt_NomArc.Text)) = 0 Then
      MsgBox "Debe ingresar la ubicación y nombre del archivo a importar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   '-----------------------
   ReDim l_arr_GenArc(0)
   r_int_FilAch = 0
   r_dbl_TotAch = 0
   r_str_FecAch = ""
   r_str_MonAch = ""
   r_str_CodMon = 0
   r_int_ValFil = 0
   r_dbl_ValTot = 0
   r_bol_ValLin = True
   r_bol_ValDNI = True
   r_str_ErrDNI = ""
   
   r_int_linea = FreeFile
   Open txt_NomArc For Input As r_int_linea
   Do While Not EOF(r_int_linea)
      Line Input #r_int_linea, r_str_Cadena
      DoEvents
      If Left(r_str_Cadena, 2) = "01" Then
         r_str_FecAch = Mid(r_str_Cadena, 20, 8)
         r_str_MonAch = Mid(r_str_Cadena, 17, 3)
         r_str_CodMon = IIf(Mid(r_str_Cadena, 17, 3) = "PEN", 1, 2)
      End If
      If Left(r_str_Cadena, 2) = "02" Then
         r_int_FilAch = r_int_FilAch + 1
         ReDim Preserve l_arr_GenArc(UBound(l_arr_GenArc) + 1)
         l_arr_GenArc(UBound(l_arr_GenArc)).r_str_TipDoc = 0
         
         r_dbl_Import = 0
         r_str_TipDoc = ""
         r_str_NumDoc = ""
         Call fs_BusNumDoc(r_str_Cadena, r_str_TipDoc, r_str_NumDoc)
         Call fs_BusImport(r_str_Cadena, r_dbl_Import)
         r_dbl_TotAch = r_dbl_TotAch + r_dbl_Import
         
         l_arr_GenArc(UBound(l_arr_GenArc)).r_str_CodMon = r_str_CodMon
         l_arr_GenArc(UBound(l_arr_GenArc)).r_str_TipDoc = r_str_TipDoc
         l_arr_GenArc(UBound(l_arr_GenArc)).r_str_NumDoc = r_str_NumDoc
         l_arr_GenArc(UBound(l_arr_GenArc)).r_str_NumOpe = Trim(Mid(r_str_Cadena, 46, 20))
         l_arr_GenArc(UBound(l_arr_GenArc)).r_dbl_Import = r_dbl_Import
         l_arr_GenArc(UBound(l_arr_GenArc)).r_str_NomCli = Mid(r_str_Cadena, 3, 33)
         l_arr_GenArc(UBound(l_arr_GenArc)).r_int_NroCta = Mid(r_str_Cadena, 68, 3)
         l_arr_GenArc(UBound(l_arr_GenArc)).r_str_TipPag = Mid(r_str_Cadena, 66, 2)
         
         If (CInt(r_str_TipDoc) <> 0 And Len(r_str_NumDoc) > 7) Then
             Dim aux As String
             aux = moddat_gf_Buscar_NomCli(CInt(r_str_TipDoc), r_str_NumDoc)
             If (Len(Trim(aux)) = 0) Then
                 r_bol_ValDNI = False
                 If r_str_ErrDNI = "" Then
                    r_str_ErrDNI = r_str_ErrDNI & CInt(r_str_TipDoc) & "-" & r_str_NumDoc
                 Else
                    r_str_ErrDNI = r_str_ErrDNI & ", " & CInt(r_str_TipDoc) & "-" & r_str_NumDoc
                 End If
             End If
         End If
      End If
      If Left(r_str_Cadena, 2) = "03" Then
         r_int_ValFil = CInt(Mid(r_str_Cadena, 3, 9))
         r_dbl_ValTot = CDbl(Mid(r_str_Cadena, 12, 15)) / 100
      End If
      If Len(r_str_Cadena) <> 152 Then
         r_bol_ValLin = False
      End If
   Loop
   Close #r_int_linea
   DoEvents
   
   If r_bol_ValDNI = False Then
      MsgBox "Los siguientes clientes no se encuentra registrado en la base de datos." & vbCrLf & _
              "Nro Documento: " & r_str_ErrDNI, vbExclamation, modgen_g_str_NomPlt
      ReDim l_arr_GenArc(0)
      Call gs_SetFocus(cmd_BuscaArc)
      Screen.MousePointer = 0
      Exit Sub
   End If
   If r_bol_ValLin = False Then
      MsgBox "La longitud de las lineas del archivo no concuerda con el formato.", vbExclamation, modgen_g_str_NomPlt
      ReDim l_arr_GenArc(0)
      Call gs_SetFocus(cmd_BuscaArc)
      Screen.MousePointer = 0
      Exit Sub
   End If
   If r_int_ValFil <> r_int_FilAch Then
      MsgBox "El nro de filas del archivo de texto no es igual al resumen en la ultima fila.", vbExclamation, modgen_g_str_NomPlt
      ReDim l_arr_GenArc(0)
      Call gs_SetFocus(cmd_BuscaArc)
      Screen.MousePointer = 0
      Exit Sub
   End If
   If Format(r_dbl_ValTot, "###,###,##0.00") <> Format(r_dbl_TotAch, "###,###,##0.00") Then
      MsgBox "La suma total de la ultima fila no es igual a la suma de todos los registros.", vbExclamation, modgen_g_str_NomPlt
      ReDim l_arr_GenArc(0)
      Call gs_SetFocus(cmd_BuscaArc)
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   pnl_FilAch.Caption = CStr(r_int_FilAch)
   pnl_TotAch.Caption = Format(r_dbl_TotAch, "###,###,##0.00") & " "
   pnl_FecAch.Caption = gf_FormatoFecha(Trim(r_str_FecAch))
   pnl_MonAch.Caption = Trim(r_str_MonAch)
         
   cmd_Grabar.Enabled = True
   Call gs_SetFocus(cmb_TipDoc)
   Screen.MousePointer = 0
End Sub

Public Function moddat_gf_Buscar_NomCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, Optional ByVal p_FlgAll As Integer) As String
   moddat_gf_Buscar_NomCli = ""
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & Trim(p_NumDoc) & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      If p_FlgAll = 1 Then
         If Len(Trim(g_rst_Listas!DatGen_ApeCas)) > 0 Then
            moddat_gf_Buscar_NomCli = Trim(g_rst_Listas!DatGen_ApePat) & " " & Trim(g_rst_Listas!DatGen_ApeMat) & " DE " & Trim(g_rst_Listas!DatGen_ApeCas) & " " & Trim(g_rst_Listas!DatGen_Nombre)
         Else
            moddat_gf_Buscar_NomCli = Trim(g_rst_Listas!DatGen_ApePat) & " " & Trim(g_rst_Listas!DatGen_ApeMat) & " " & Trim(g_rst_Listas!DatGen_Nombre)
         End If
      Else
         moddat_gf_Buscar_NomCli = Trim(g_rst_Listas!DatGen_ApePat) & " " & Trim(g_rst_Listas!DatGen_ApeMat) & " " & Trim(g_rst_Listas!DatGen_Nombre)
      End If
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub fs_BusNumDoc(p_Cadena As String, ByRef p_TipDoc As String, ByRef p_NumDoc As String)
Dim r_str_DocIde   As String

   r_str_DocIde = Format(Trim(Mid(p_Cadena, 33, 13)), "#############")
   p_TipDoc = Trim(Mid(p_Cadena, 37, 1))
   p_NumDoc = Trim(Mid(p_Cadena, 38, 8))
   
   If (p_TipDoc = 0) Then
       p_TipDoc = Trim(Mid(p_Cadena, 36, 1))
       p_NumDoc = Trim(Mid(p_Cadena, 37, 9))
       If (p_TipDoc = 0) Then
           p_TipDoc = Trim(Mid(p_Cadena, 35, 1))
           p_NumDoc = Trim(Mid(p_Cadena, 36, 10))
           If (p_TipDoc = 0) Then
               p_TipDoc = Trim(Mid(p_Cadena, 34, 1))
               p_NumDoc = Trim(Mid(p_Cadena, 35, 11))
               If (p_TipDoc = 0) Then
                   p_TipDoc = Trim(Mid(p_Cadena, 33, 1))
                   p_NumDoc = Trim(Mid(p_Cadena, 34, 12))
               End If
           End If
       End If
   End If
End Sub

Private Sub fs_BusImport(p_Cadena As String, p_Import As Double)
   p_Import = 0
   p_Import = CDbl(Mid(p_Cadena, 96, 13) & "." & Mid(p_Cadena, 109, 2))
End Sub

Private Sub fs_CargarPrv()
   ReDim l_arr_MaePrv(0)
   cmb_Proveedor.Clear
   cmb_Proveedor.Text = ""
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   
   If (cmb_TipDoc.ListIndex = -1) Then
       Exit Sub
   End If
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC, A.MAEPRV_CODSIC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   If moddat_g_int_FlgGrb = 1 Then 'INSERT
      g_str_Parame = g_str_Parame & " AND A.MAEPRV_SITUAC = 1 "
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY A.MAEPRV_RAZSOC ASC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      cmb_Proveedor.AddItem Trim(g_rst_Genera!MAEPRV_NUMDOC & "") & " - " & Trim(g_rst_Genera!MaePrv_RazSoc & "")
      
      ReDim Preserve l_arr_MaePrv(UBound(l_arr_MaePrv) + 1)
      l_arr_MaePrv(UBound(l_arr_MaePrv)).Genera_Codigo = Trim(g_rst_Genera!MAEPRV_NUMDOC & "")
      l_arr_MaePrv(UBound(l_arr_MaePrv)).Genera_Nombre = Trim(g_rst_Genera!MaePrv_RazSoc & "")
      l_arr_MaePrv(UBound(l_arr_MaePrv)).Genera_Prefij = Trim(g_rst_Genera!MAEPRV_CODSIC & "")
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub fs_Buscar_prov()
Dim r_str_NumDoc As String

   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   r_str_NumDoc = ""
   
   If (moddat_g_int_FlgGrb = 1) Then
       If cmb_TipDoc.ListIndex = -1 Then
          MsgBox "Debe seleccionar el tipo de documento de identidad.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_TipDoc)
          Exit Sub
       End If
       If cmb_Proveedor.ListIndex = -1 Then
          Exit Sub
       End If
      
       If (fs_ValNumDoc() = False) Then
           Exit Sub
       End If
   End If
   
   r_str_NumDoc = fs_NumDoc(cmb_Proveedor.Text)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_CODBNC_MN1, A.MAEPRV_CTACRR_MN1, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN1, A.MAEPRV_CODBNC_MN2, A.MAEPRV_CTACRR_MN2, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN2, A.MAEPRV_CODBNC_MN3, A.MAEPRV_CTACRR_MN3, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN3, A.MAEPRV_CODBNC_DL1, A.MAEPRV_CTACRR_DL1, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL1, A.MAEPRV_CODBNC_DL2, A.MAEPRV_CTACRR_DL2, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL2, A.MAEPRV_CODBNC_DL3, A.MAEPRV_CTACRR_DL3, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL3, A.MAEPRV_CONDIC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   If (moddat_g_int_FlgGrb = 1 Or moddat_g_int_FlgGrb = 0) Then
       g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
       g_str_Parame = g_str_Parame & "    AND TRIM(A.MAEPRV_NUMDOC) = '" & Trim(r_str_NumDoc) & "' "
   Else
       g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & moddat_g_str_TipDoc
       g_str_Parame = g_str_Parame & "    AND TRIM(A.MAEPRV_NUMDOC) = '" & Trim(moddat_g_str_NumDoc) & "' "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      MsgBox "No se ha encontrado el proveedor.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Proveedor)
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Sub
   End If
   
   If (moddat_g_int_FlgGrb = 1) Then
       If (g_rst_GenAux!MAEPRV_CONDIC = 2) Then
          MsgBox "El proveedor se encuentra en condición de NO HABIDO, revisar sunat.", vbExclamation, modgen_g_str_NomPlt
          g_rst_GenAux.Close
          Set g_rst_GenAux = Nothing
          Exit Sub
       End If
       'Call gs_SetFocus(txt_Descrip)
   End If
      
   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)

   If (g_rst_GenAux!MAEPRV_CODBNC_MN1 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN1, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN1 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN1 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_MN2 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN2)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN2, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN2 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN2 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_MN3 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN3)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN3, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN3 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN3 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   
   If (g_rst_GenAux!MAEPRV_CODBNC_DL1 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL1, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL1 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL1 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_DL2 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL2)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL2, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL2 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL2 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_DL3 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL3)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL3, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL3 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL3 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   
   Call fs_CargarBancos
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Sub

Private Sub fs_CargarBancos()
Dim r_bol_Estado   As Boolean
Dim r_int_File     As Integer
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   
   If (cmb_Moneda.ListIndex = -1) Then
       Exit Sub
   End If
   
   'soles
   If (cmb_Moneda.ListIndex = 0) Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteSol)
           r_bol_Estado = True
           For r_int_File = 0 To cmb_Banco.ListCount - 1
               If (Trim(cmb_Banco.ItemData(r_int_File)) = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Codigo)) Then
                   r_bol_Estado = False
                   Exit For
               End If
           Next
           If (r_bol_Estado = True) Then
               cmb_Banco.AddItem Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Nombre)
               cmb_Banco.ItemData(cmb_Banco.NewIndex) = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Codigo)
           End If
       Next
   End If
   'dolares
   If (cmb_Moneda.ListIndex = 1) Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteDol)
           r_bol_Estado = True
           For r_int_File = 0 To cmb_Banco.ListCount - 1
               If (Trim(cmb_Banco.ItemData(r_int_File)) = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Codigo)) Then
                   r_bol_Estado = False
                   Exit For
               End If
           Next
           If (r_bol_Estado = True) Then
               cmb_Banco.AddItem Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Nombre)
               cmb_Banco.ItemData(cmb_Banco.NewIndex) = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Codigo)
           End If
       Next
   End If
End Sub

Private Function fs_ValNumDoc() As Boolean
Dim r_str_NumDoc As String
   fs_ValNumDoc = True
   r_str_NumDoc = ""

   r_str_NumDoc = fs_NumDoc(cmb_Proveedor.Text)
   If (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1) Then 'DNI - 8
       If Len(Trim(r_str_NumDoc)) <> 8 Then
          MsgBox "El documento de identidad es de 8 digitos.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_Proveedor)
          fs_ValNumDoc = False
       End If
   ElseIf (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6) Then 'RUC - 11
       If Not gf_Valida_RUC(Trim(r_str_NumDoc), Mid(Trim(r_str_NumDoc), Len(Trim(r_str_NumDoc)), 1)) Then
          MsgBox "El Número de RUC no es valido.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_Proveedor)
          fs_ValNumDoc = False
       End If
   Else 'OTROS
       If Len(Trim(cmb_Proveedor.Text)) = 0 Then
          MsgBox "Debe ingresar un numero de documento.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_Proveedor)
          fs_ValNumDoc = False
       End If
   End If
End Function

Private Function fs_NumDoc(p_Cadena As String) As String
   fs_NumDoc = ""
   If (cmb_TipDoc.ListIndex > -1) Then
      If (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1) Then
          fs_NumDoc = Mid(p_Cadena, 1, 8)
      ElseIf (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6) Then
          fs_NumDoc = Mid(p_Cadena, 1, 11)
      Else
          If p_Cadena <> "" Then
             fs_NumDoc = Trim(Mid(p_Cadena, 1, InStr(Trim(p_Cadena), "-") - 1))
          End If
      End If
   End If
End Function

Private Sub cmb_Banco_Click()
Dim r_str_Cadena  As String
   
   cmb_CtaCte.Clear
   r_str_Cadena = ""
   lbl_Cuenta.Caption = "Cuenta:"
   
   If (cmb_Moneda.ListIndex = -1) Then
       Exit Sub
   End If
   
   If (cmb_Moneda.ListIndex = 0) Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteSol)
           r_str_Cadena = ""
           If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Codigo)) Then
               If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = 11) Then 'Banco continental
                   r_str_Cadena = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Prefij)
                   lbl_Cuenta.Caption = "Cuenta Corriente:"
               Else
                   r_str_Cadena = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Refere)
                   lbl_Cuenta.Caption = "CCI:"
               End If
           End If
           If (Len(Trim(r_str_Cadena)) > 0) Then
               cmb_CtaCte.AddItem Trim(r_str_Cadena)
           End If
       Next
   End If
   
   If (cmb_Moneda.ListIndex = 1) Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteDol)
           r_str_Cadena = ""
           If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Codigo)) Then
               If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = 11) Then 'Banco continental
                   r_str_Cadena = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Prefij)
                   lbl_Cuenta.Caption = "Cuenta Corriente:"
               Else
                   r_str_Cadena = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Refere)
                   lbl_Cuenta.Caption = "CCI:"
               End If
           End If
           If (Len(Trim(r_str_Cadena)) > 0) Then
               cmb_CtaCte.AddItem Trim(r_str_Cadena)
           End If
       Next
   End If
End Sub

Private Sub fs_Guardar_Log(p_NomFil As String, ByRef p_FecPro As String, p_NumPro As Integer)
Dim r_int_NumFil     As Integer
Dim r_str_Cadena     As String
Dim r_str_CodBco     As String
Dim r_str_NumCta     As String
Dim r_str_FecRec     As String
Dim r_str_NumOpe     As String
Dim r_str_FecPag     As String
Dim r_str_NumDoc     As String
Dim r_int_TipDoc     As Integer
Dim r_int_TipMon     As Integer
Dim r_int_TipPag     As Integer
Dim r_int_NumCuo     As Integer
Dim r_dbl_ImpDep     As Double
Dim r_str_CadErr     As String
Dim r_int_CntErr     As Integer
Dim r_int_FlgErr     As Integer
Dim r_str_DocIde     As String
Dim r_str_HorIni     As String
Dim r_str_HorFin     As String
Dim r_str_Situac     As String
Dim r_str_NomCli     As String
Dim r_int_NumReg     As Integer
Dim r_dbl_ImpTot     As Double
Dim r_int_ConErr     As Integer
Dim r_int_SinErr     As Integer
Dim r_int_NumPro     As Integer
Dim r_int_ErrorUpd   As Integer
Dim r_int_NFila      As Integer

   p_FecPro = "" 'PK
   p_NumPro = 0  'PK
   r_str_CodBco = "000002" 'l_str_CodBan
   r_str_NumCta = ""
   r_int_TipMon = 0
   r_str_FecRec = ""
   r_dbl_ImpTot = 0
   r_int_SinErr = 0
   r_int_ConErr = 0
   r_int_ErrorUpd = 0
   r_str_HorIni = Format(Time, "hhmmss")
   DoEvents
   
   r_int_NumFil = FreeFile
   Open p_NomFil For Input As r_int_NumFil
   
   'Leyendo Cabecera del Archivo
   Line Input #r_int_NumFil, r_str_Cadena
   
   If Left(r_str_Cadena, 2) = "01" Then
      r_str_NumCta = Mid(r_str_Cadena, 28, 18)
      r_str_FecRec = Mid(r_str_Cadena, 20, 8)
      
      Select Case Mid(r_str_Cadena, 17, 3)
         Case "USD": r_int_TipMon = 2
         Case "PEN": r_int_TipMon = 1
      End Select
   End If
   
   '*** inicializa log
   moddat_g_int_CntErr = 0
   g_str_Parame = "USP_CRE_PROPAGCAB ("
   g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 1 & ", "
   g_str_Parame = g_str_Parame & "'" & p_NomFil & "', "
   g_str_Parame = g_str_Parame & "'" & Dir(p_NomFil, vbArchive) & "', "
   g_str_Parame = g_str_Parame & r_str_FecRec & ", "
   g_str_Parame = g_str_Parame & CInt(r_str_CodBco) & ", "
   g_str_Parame = g_str_Parame & "'" & r_str_NumCta & "', "
   g_str_Parame = g_str_Parame & CInt(r_int_TipMon) & " , "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & 0 & ", "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & 1 & " ) "
   
   Do While (moddat_g_int_CntErr = 0)
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            moddat_g_int_CntErr = 1
            Close #r_int_NumFil
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      Else
         moddat_g_int_CntErr = 1
      End If
   Loop
      
   g_rst_Princi.MoveFirst
   r_int_NumPro = g_rst_Princi!CORRELATIVO
   
   r_int_FlgErr = 0
   r_int_NFila = 0
   
   Do While Not EOF(r_int_NumFil)
      Line Input #r_int_NumFil, r_str_Cadena
      DoEvents
      
      r_int_CntErr = 0
      r_int_ErrorUpd = 0
      
      If Left(r_str_Cadena, 2) = "02" Then
         r_str_NomCli = Trim(Mid(r_str_Cadena, 3, 30))
         r_str_DocIde = Format(Trim(Mid(r_str_Cadena, 33, 13)), "#############")
         r_int_TipDoc = CInt(Mid(r_str_DocIde, 1, 1))
         r_str_NumDoc = Trim(Mid(r_str_DocIde, 2))
         r_str_NumOpe = Trim(Mid(r_str_Cadena, 46, 20))
         r_int_TipPag = CInt(Mid(r_str_Cadena, 66, 2))
         r_int_NumCuo = CInt(Mid(r_str_Cadena, 68, 3))
         r_str_FecPag = Mid(r_str_Cadena, 136, 8)
         r_dbl_ImpDep = CDbl(Mid(r_str_Cadena, 96, 13) & "." & Mid(r_str_Cadena, 109, 2))
         
         r_int_NFila = r_int_NFila + 1
         r_str_Situac = 1
         
         '*** actualiza log
         g_str_Parame = "USP_CRE_PROPAGDET ("
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & r_int_NumPro & ", "
         g_str_Parame = g_str_Parame & r_int_NFila & ", "
         g_str_Parame = g_str_Parame & r_str_Situac & " , "
         g_str_Parame = g_str_Parame & "'" & r_str_NumOpe & "' , "
         g_str_Parame = g_str_Parame & "'" & r_str_NumDoc & "' , "
         g_str_Parame = g_str_Parame & "'" & r_str_NomCli & "', "
         g_str_Parame = g_str_Parame & r_int_TipPag & ", "
         g_str_Parame = g_str_Parame & r_str_FecPag & ", "
         g_str_Parame = g_str_Parame & r_dbl_ImpDep & ", "
         g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
         g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
         
         Do While (r_int_ErrorUpd = 0)
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  r_int_ErrorUpd = 1
               Else
                  r_int_ErrorUpd = 0
               End If
            Else
               r_int_ErrorUpd = 1
            End If
         Loop
         
      End If
   Loop
   
   Close #r_int_NumFil
   
   r_str_HorFin = Format(Time, "hhmmss")
   
   If Left(r_str_Cadena, 2) = "03" Then
      r_int_NumReg = CInt(Mid(r_str_Cadena, 5, 7))
      r_dbl_ImpTot = CDbl(Mid(r_str_Cadena, 12, 15)) / 100
      
      '*** finaliza log
      g_str_Parame = "USP_CRE_PROPAGCAB ("
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & r_int_NumPro & ", "
      g_str_Parame = g_str_Parame & 2 & ", "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & r_str_FecRec & ", "
      g_str_Parame = g_str_Parame & CInt(r_str_CodBco) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_NumCta & "', "
      g_str_Parame = g_str_Parame & CInt(r_int_TipMon) & " , "
      g_str_Parame = g_str_Parame & r_int_NumReg & ", "
      g_str_Parame = g_str_Parame & r_dbl_ImpTot & ", "
      g_str_Parame = g_str_Parame & r_int_NumFil & ", "
      g_str_Parame = g_str_Parame & r_int_ConErr & ", "
      g_str_Parame = g_str_Parame & r_int_SinErr & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & 0 & " ) "
      
      Do While (r_int_ErrorUpd = 0)
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               r_int_ErrorUpd = 1
            Else
               r_int_ErrorUpd = 0
            End If
         Else
            r_int_ErrorUpd = 1
         End If
      Loop
   End If
   
   'DEVOLVER CODIGO PK
   p_FecPro = Format(CDate(moddat_g_str_FecSis), "yyyymmdd") 'PK
   p_NumPro = r_int_NumPro  'PK
End Sub

Private Sub fs_Cargar_Datos()
Dim r_int_Contad As Integer

   Call gs_SetFocus(cmb_TipDoc)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.PROPAG_CODPAG, A.PROPAG_FECREG, A.PROPAG_TIPDOC, A.PROPAG_NUMDOC,  "
   g_str_Parame = g_str_Parame & "        A.PROPAG_CODMON, A.PROPAG_CODBCO, A.PROPAG_CTACRR, PROPAG_DESCRP,  "
   g_str_Parame = g_str_Parame & "        DECODE(B.PAGCAB_MONEDA,1,'PEN','USD') REC_MONEDA, B.PAGCAB_NUMREGFIL, B.PAGCAB_FECREC,  "
   g_str_Parame = g_str_Parame & "        B.PAGCAB_TOTPAGFIL, B.PAGCAB_FECPRO, TRIM(B.PAGCAB_RUTFIL) || TRIM(B.PAGCAB_NOMFIL) AS ARCH_RUTA  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_PROPAG_ARH A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CRE_PROPAGCAB B ON B.PAGCAB_FECPRO = A.PROPAG_FECPRO AND B.PAGCAB_NUMPRO = A.PROPAG_NUMPRO  "
   g_str_Parame = g_str_Parame & "  WHERE A.PROPAG_CODPAG = " & CLng(moddat_g_str_Codigo)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_Codigo.Caption = Format(g_rst_Princi!PROPAG_CODPAG, "0000000000")
      pnl_FecProc.Caption = gf_FormatoFecha(g_rst_Princi!PROPAG_FECREG)
      txt_NomArc.Text = g_rst_Princi!ARCH_RUTA
      pnl_FecAch.Caption = gf_FormatoFecha(g_rst_Princi!PAGCAB_FECREC)
      pnl_MonAch.Caption = g_rst_Princi!REC_MONEDA
      pnl_FilAch.Caption = CStr(g_rst_Princi!PAGCAB_NUMREGFIL)
      pnl_TotAch.Caption = Format(g_rst_Princi!PAGCAB_TOTPAGFIL, "###,###,###,##0.00") & " "
      
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!PROPAG_TIPDOC)
      cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, g_rst_Princi!PROPAG_NUMDOC & "", 0)
      txt_Descrip.Text = Trim(g_rst_Princi!PROPAG_DESCRP & "")
      Call gs_BuscarCombo_Item(cmb_Moneda, g_rst_Princi!PROPAG_CODMON)
      Call gs_BuscarCombo_Item(cmb_Banco, g_rst_Princi!PROPAG_CODBCO)
      
      Call gs_BuscarCombo_Text(cmb_CtaCte, g_rst_Princi!PROPAG_CTACRR, -1)
      'cmb_CtaCte.Text = g_rst_Princi!PROPAG_CTACRR
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Function fs_ComboIndex(p_Combo As ComboBox, Cadena As String, p_Tipo As Integer) As Integer
Dim r_int_Contad As Integer

   fs_ComboIndex = -1
   For r_int_Contad = 0 To p_Combo.ListCount - 1
       If Trim(Cadena) = Trim(Mid(p_Combo.List(r_int_Contad), 1, InStr(Trim(p_Combo.List(r_int_Contad)), "-") - 1)) Then
          fs_ComboIndex = r_int_Contad
          Exit For
       End If
   Next
End Function

Private Sub fs_GeneraAsiento(p_CodPag As String, p_AsiGen As String)
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_int_NumAsi        As Integer
Dim r_str_Glosa         As String
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_dbl_TotSol        As Double
Dim r_dbl_TotDol        As Double
Dim r_str_DebHab        As String
Dim r_dbl_TipSbs        As Double
Dim r_str_CadAux        As String
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer
Dim r_int_NumIte        As Integer
Dim r_str_FecPrPgoC     As String
Dim r_str_FecPrPgoL     As String
Dim r_str_CtaDeb        As String
Dim r_str_CtaHab        As String
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "O"
   r_int_NumLib = 6
   p_AsiGen = ""
   r_str_CadAux = ""
             
   'Inicializa variables
   r_int_NumAsi = 0
   r_int_NumIte = 0
   r_str_FecPrPgoC = Format(Trim(pnl_FecProc.Caption), "yyyymmdd")
   r_str_FecPrPgoL = Format(Trim(pnl_FecProc.Caption), "dd/mm/yyyy")
   
   'TipCam = 1 - Comercial / 2 - SBS / 3 - Sunat / 4 - BCR
   'TipTip = 1 - Venta / 2 - Compra
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, r_str_FecPrPgoC, 1)
   
   If l_int_TipBnc = 1 Then
      r_str_Glosa = "RECAUDO " & Format(Trim(pnl_FecAch.Caption), "ddmmyyyy") & "/BANCO GNB"
   Else
      r_str_Glosa = "RECAUDO " & Format(Trim(pnl_FecAch.Caption), "ddmmyyyy") & "/BANCO COMERCIO"
   End If
   
   'r_int_PerMes = modctb_int_PerMes 'Format(r_str_FecPrPgoL, "mm")
   'r_int_PerAno = modctb_int_PerAno 'Format(r_str_FecPrPgoL, "yyyy")
   r_int_PerMes = Format(r_str_FecPrPgoL, "mm")
   r_int_PerAno = Format(r_str_FecPrPgoL, "yyyy")
   
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
   p_AsiGen = CStr(r_int_NumAsi)
   
   'Insertar en CABECERA
   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
        r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
        
   r_str_CtaDeb = "291807010103"
   r_str_CtaHab = "251419010109"
   r_int_NumIte = 0
   r_dbl_TotSol = 0
   r_dbl_TotDol = 0
   r_str_Glosa = ""
   For l_int_Contar = 1 To UBound(l_arr_GenArc)
       r_str_Glosa = "RECAUDO " & Format(Trim(pnl_FecAch.Caption), "ddmmyyyy") & "/" & l_arr_GenArc(l_int_Contar).r_str_NumDoc & "/" & l_arr_GenArc(l_int_Contar).r_str_NumOpe
       
       r_dbl_MtoSol = 0
       r_dbl_MtoDol = 0
       If l_arr_GenArc(l_int_Contar).r_str_CodMon = 1 Then     'SOLES
          r_dbl_MtoSol = l_arr_GenArc(l_int_Contar).r_dbl_Import
          r_dbl_MtoDol = Format(CDbl(l_arr_GenArc(l_int_Contar).r_dbl_Import / r_dbl_TipSbs), "###,###,##0.00")  'Importe / CONVERTIDO
       ElseIf l_arr_GenArc(l_int_Contar).r_str_CodMon = 2 Then 'DOLARES
          r_dbl_MtoSol = Format(CDbl(l_arr_GenArc(l_int_Contar).r_dbl_Import * r_dbl_TipSbs), "###,###,##0.00")  'Importe * CONVERTIDO
          r_dbl_MtoDol = l_arr_GenArc(l_int_Contar).r_dbl_Import
       End If
       
       r_dbl_TotSol = r_dbl_TotSol + r_dbl_MtoSol
       r_dbl_TotDol = r_dbl_TotDol + r_dbl_MtoDol
       r_int_NumIte = r_int_NumIte + 1
       Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                            r_int_NumAsi, r_int_NumIte, r_str_CtaDeb, CDate(r_str_FecPrPgoL), _
                                            r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
   Next
   
   If l_int_TipBnc = 1 Then
      r_str_Glosa = "RECAUDO " & Format(Trim(pnl_FecAch.Caption), "ddmmyyyy") & "/" & fs_NumDoc(cmb_Proveedor.Text) & "/BANCO GNB"
   Else
      r_str_Glosa = "RECAUDO " & Format(Trim(pnl_FecAch.Caption), "ddmmyyyy") & "/" & fs_NumDoc(cmb_Proveedor.Text) & "/BANCO COMERCIO"
   End If
   
   r_int_NumIte = r_int_NumIte + 1
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, r_int_NumIte, r_str_CtaHab, CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "H", r_dbl_TotSol, r_dbl_TotDol, 1, CDate(r_str_FecPrPgoL))
        
   'Actualiza flag de contabilizacion
   r_str_CadAux = ""
   r_str_CadAux = r_str_Origen & "/" & r_int_PerAno & "/" & Format(r_int_PerMes, "00") & "/" & Format(r_int_NumLib, "00") & "/" & r_int_NumAsi
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " UPDATE CNTBL_PROPAG_ARH  "
   g_str_Parame = g_str_Parame & "    SET PROPAG_FLGCTB = 1,  "
   g_str_Parame = g_str_Parame & "        PROPAG_FECCTB = " & Format(moddat_g_str_FecSis, "yyyymmdd") & ",  "
   g_str_Parame = g_str_Parame & "        PROPAG_DATCTB = '" & r_str_CadAux & "',  "
   g_str_Parame = g_str_Parame & "        PROPAG_TIPCAM = " & r_dbl_TipSbs
   g_str_Parame = g_str_Parame & "  WHERE PROPAG_CODPAG  = " & CLng(p_CodPag)
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
   
   'Enviar a la tabla de autorizaciones
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT ( "
   g_str_Parame = g_str_Parame & " " & CLng(p_CodPag) & ", " 'COMAUT_CODOPE
   g_str_Parame = g_str_Parame & " " & Format(pnl_FecProc.Caption, "yyyymmdd") & ", " 'COMAUT_FECOPE
   g_str_Parame = g_str_Parame & " " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", "      'COMAUT_TIPDOC
   g_str_Parame = g_str_Parame & " '" & fs_NumDoc(cmb_Proveedor.Text) & "', "    'COMAUT_NUMDOC
   g_str_Parame = g_str_Parame & " " & cmb_Moneda.ItemData(cmb_Moneda.ListIndex) & ", "      'COMAUT_CODMON
   g_str_Parame = g_str_Parame & " " & CDbl(pnl_TotAch.Caption) & ", " 'COMAUT_IMPPAG
   g_str_Parame = g_str_Parame & " " & cmb_Banco.ItemData(cmb_Banco.ListIndex) & ", "  'COMAUT_CODBNC
   g_str_Parame = g_str_Parame & " '" & Trim(cmb_CtaCte.Text) & "', " 'COMAUT_CTACRR
   g_str_Parame = g_str_Parame & " '" & r_str_CtaHab & "', "  'COMAUT_CTACTB
   g_str_Parame = g_str_Parame & " '" & r_str_CadAux & "',  " 'COMAUT_DATCTB
   g_str_Parame = g_str_Parame & " '" & Left(Trim(txt_Descrip.Text), 150) & "', " 'COMAUT_DESCRP
   g_str_Parame = g_str_Parame & " 1, "  'COMAUT_TIPOPE
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "  'SEGUSUCRE
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "  'SEGPLTCRE
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "  'SEGTERCRE
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "  'SEGSUCCRE

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Proveedor)
   End If
End Sub

Private Sub cmb_Proveedor_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_Descrip)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub txt_Descrip_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Moneda)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub cmb_Banco_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_CtaCte)
   End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmb_Banco)
   End If
End Sub

Private Sub cmb_CtaCte_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

