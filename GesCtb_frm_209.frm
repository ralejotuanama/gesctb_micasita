VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_PagCom_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17025
   Icon            =   "GesCtb_frm_209.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   17025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17100
      _Version        =   65536
      _ExtentX        =   30162
      _ExtentY        =   12621
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
         TabIndex        =   1
         Top             =   60
         Width           =   16905
         _Version        =   65536
         _ExtentX        =   29810
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
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registros de Pagos Aprobados"
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
            Picture         =   "GesCtb_frm_209.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   670
         Left            =   60
         TabIndex        =   3
         Top             =   780
         Width           =   16900
         _Version        =   65536
         _ExtentX        =   29810
         _ExtentY        =   1182
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.19
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin Threed.SSFrame SSFrame1 
            Height          =   620
            Left            =   11970
            TabIndex        =   23
            Top             =   30
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   1094
            _StockProps     =   14
            Caption         =   "Calculo Descuentos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.32
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.OptionButton rb_Ninguno 
               Caption         =   "Ninguno"
               Height          =   300
               Left            =   480
               TabIndex        =   24
               Top             =   270
               Width           =   1095
            End
            Begin VB.OptionButton rb_4TA 
               Caption         =   "4ta"
               Height          =   300
               Left            =   2910
               TabIndex        =   26
               Top             =   270
               Width           =   795
            End
            Begin VB.OptionButton rb_ITF 
               Caption         =   "ITF"
               Height          =   300
               Left            =   1890
               TabIndex        =   25
               Top             =   270
               Width           =   795
            End
         End
         Begin VB.CommandButton cmb_Adicionar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_209.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_209.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   16290
            Picture         =   "GesCtb_frm_209.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_209.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5475
         Left            =   60
         TabIndex        =   6
         Top             =   1495
         Width           =   16905
         _Version        =   65536
         _ExtentX        =   29810
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   15780
            TabIndex        =   14
            Top             =   60
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
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
               Left            =   810
               TabIndex        =   15
               Top             =   20
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   3540
            TabIndex        =   8
            Top             =   60
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Evaluador"
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   2505
            TabIndex        =   9
            Top             =   60
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1870
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5055
            Left            =   30
            TabIndex        =   7
            Top             =   360
            Width           =   16845
            _ExtentX        =   29713
            _ExtentY        =   8916
            _Version        =   393216
            Rows            =   30
            Cols            =   25
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   4845
            TabIndex        =   10
            Top             =   60
            Width           =   2730
            _Version        =   65536
            _ExtentX        =   4815
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
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
         Begin Threed.SSPanel pnl_Tit_SitIns 
            Height          =   285
            Left            =   11925
            TabIndex        =   12
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1940
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Pagar"
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
            Left            =   9210
            TabIndex        =   18
            Top             =   60
            Width           =   1950
            _Version        =   65536
            _ExtentX        =   3440
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuenta Corriente"
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
            Left            =   1140
            TabIndex        =   19
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   13005
            TabIndex        =   20
            Top             =   60
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1676
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Aplicación"
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
            Left            =   13935
            TabIndex        =   21
            Top             =   60
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1411
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Dscto."
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
            Left            =   14700
            TabIndex        =   22
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1940
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Pago Neto"
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
            Left            =   7560
            TabIndex        =   27
            Top             =   60
            Width           =   1660
            _Version        =   65536
            _ExtentX        =   2928
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
         Begin Threed.SSPanel pnl_Tit_IngIns 
            Height          =   285
            Left            =   11145
            TabIndex        =   13
            Top             =   60
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1393
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
      End
   End
End
Attribute VB_Name = "frm_Ctb_PagCom_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_Contar As Integer
Dim l_int_FilAnt As Integer

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 12) = ""
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 12) = "X"
         Next r_Fila
      End If
   Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub cmb_Adicionar_Click()
Dim r_int_Contar  As Integer
Dim r_int_Fila    As Integer
Dim r_bol_Estado  As Boolean
Dim r_int_TotFil  As Integer
Dim r_str_Cadena  As String

   With frm_Ctb_PagCom_02
        r_int_TotFil = 0
        r_bol_Estado = True
        r_str_Cadena = ""
        For r_int_Fila = 0 To grd_Listad.Rows - 1
            If Trim(grd_Listad.TextMatrix(r_int_Fila, 12)) = "X" Then
               For r_int_Contar = 0 To .grd_Listad.Rows - 1
                   If Trim(.grd_Listad.TextMatrix(r_int_Contar, 0)) = Trim(grd_Listad.TextMatrix(r_int_Fila, 0)) Then
                      r_bol_Estado = False
                      r_str_Cadena = r_str_Cadena & "-" & Trim(grd_Listad.TextMatrix(r_int_Fila, 0))
                   End If
               Next
               r_int_TotFil = r_int_TotFil + 1
            End If
        Next
        
        If r_bol_Estado = False Then
           MsgBox "Los Registros " & r_str_Cadena & " ya fueron adicionados.", vbExclamation, modgen_g_str_NomPlt
           Exit Sub
        Else
           If r_int_TotFil = 0 Then
              MsgBox "No hay filas seleccionadas para adicionar.", vbExclamation, modgen_g_str_NomPlt
              Exit Sub
           End If
           
           Dim r_str_NumDoc As String
           Dim r_str_NomPrv As String
           r_str_NumDoc = ""
           r_str_NomPrv = ""
           'validar el tipo de pago
           For r_int_Fila = 0 To grd_Listad.Rows - 1
               If Trim(grd_Listad.TextMatrix(r_int_Fila, 12)) = "X" Then
                  If .cmb_TipPag.ItemData(.cmb_TipPag.ListIndex) = 1 Or .cmb_TipPag.ItemData(.cmb_TipPag.ListIndex) = 6 Or .cmb_TipPag.ItemData(.cmb_TipPag.ListIndex) = 8 Then
                     'TRANSFERENCIA(1), PAGO PROVEEDORES(6), HABERES(8)
                     If Trim(grd_Listad.TextMatrix(r_int_Fila, 6) & "") = "" Then
                        MsgBox "El tipo pago seleccionado obliga a los registros a adicionar tengan cuenta corriente.", vbExclamation, modgen_g_str_NomPlt
                        Exit Sub
                     End If
                     'TODO MENOS DETRACCION
                     If Mid(Trim(grd_Listad.TextMatrix(r_int_Fila, 0)), 1, 2) = "06" And CInt(grd_Listad.TextMatrix(r_int_Fila, 21)) = 2 Then
                        MsgBox "No se puede adicionar registros de tipo detracción, por el tipo de pago seleccionado.", vbExclamation, modgen_g_str_NomPlt
                        Exit Sub
                     End If
                     If .cmb_TipPag.ItemData(.cmb_TipPag.ListIndex) = 8 Then
                        'HABERES (8)
                        If Trim(grd_Listad.TextMatrix(r_int_Fila, 23)) = "" Then
                           MsgBox "El registro " & grd_Listad.TextMatrix(r_int_Fila, 0) & " no tiene código de planilla.", vbExclamation, modgen_g_str_NomPlt
                           Exit Sub
                        End If
                        If Trim(grd_Listad.TextMatrix(r_int_Fila, 24) & "") <> "2" Then
                           MsgBox "El registro " & grd_Listad.TextMatrix(r_int_Fila, 0) & " no es un personal interno.", vbExclamation, modgen_g_str_NomPlt
                           Exit Sub
                        End If
                        'monedas distintas
                        If CLng(.cmb_Moneda.ItemData(.cmb_Moneda.ListIndex)) <> CLng(grd_Listad.TextMatrix(r_int_Fila, 17)) Then
                           MsgBox "El tipo moneda seleccionada obliga a los registros a adicionar sean en ." & Trim(.cmb_Moneda.Text), vbExclamation, modgen_g_str_NomPlt
                           Exit Sub
                        End If
                     End If
                  ElseIf .cmb_TipPag.ItemData(.cmb_TipPag.ListIndex) = 4 Then
                     'DETRACCION
                     If Trim(grd_Listad.TextMatrix(r_int_Fila, 6) & "") = "" Then
                        MsgBox "El tipo pago seleccionado obliga a los registros a adicionar tengan cuenta corriente.", vbExclamation, modgen_g_str_NomPlt
                        Exit Sub
                     End If
                     If Left(Trim(grd_Listad.TextMatrix(r_int_Fila, 0)), 2) <> "06" Then
                        MsgBox "Para el pago de detracción, el origen de los registros deben de ser: registro de compras", vbExclamation, modgen_g_str_NomPlt
                        Exit Sub
                     End If
                     If CInt(grd_Listad.TextMatrix(r_int_Fila, 21)) <> 2 Then
                        MsgBox "Solo se pueden adicionar registro de tipo detracción, por el tipo de pago seleccionado.", vbExclamation, modgen_g_str_NomPlt
                        Exit Sub
                     End If
                     If CInt(grd_Listad.TextMatrix(r_int_Fila, 17)) <> .cmb_Moneda.ItemData(.cmb_Moneda.ListIndex) Then '--solo soles
                        MsgBox "Solo se admiten registros en soles para los pagos en detracción.", vbExclamation, modgen_g_str_NomPlt
                        Exit Sub
                     End If
                  Else
                     'CHEQUE O CARTA
                     If CLng(.cmb_Moneda.ItemData(.cmb_Moneda.ListIndex)) <> CLng(grd_Listad.TextMatrix(r_int_Fila, 17)) Then
                        MsgBox "El tipo moneda seleccionada obliga a los registros a adicionar sean en ." & Trim(.cmb_Moneda.Text), vbExclamation, modgen_g_str_NomPlt
                        Exit Sub
                     End If
                     'TODO MENOS DETRACCION
                     If Mid(Trim(grd_Listad.TextMatrix(r_int_Fila, 0)), 1, 2) = "06" And CInt(grd_Listad.TextMatrix(r_int_Fila, 21)) = 2 Then
                        MsgBox "No se puede adicionar registros de tipo detracción, por el tipo de pago seleccionado.", vbExclamation, modgen_g_str_NomPlt
                        Exit Sub
                     End If
                     'mismo proveedor
                     If .grd_Listad.Rows = 0 Then
                        If r_str_NumDoc = "" Then
                           r_str_NumDoc = Trim(grd_Listad.TextMatrix(r_int_Fila, 22))
                           r_str_NomPrv = Trim(grd_Listad.TextMatrix(r_int_Fila, 4))
                         End If
                     Else
                        r_str_NumDoc = Trim(.grd_Listad.TextMatrix(0, 3))
                        r_str_NomPrv = Trim(.grd_Listad.TextMatrix(0, 4))
                     End If
                     If r_str_NumDoc <> Trim(grd_Listad.TextMatrix(r_int_Fila, 22)) Then
                        MsgBox "Solo se pueden adicionar los registros del proveedor:" & vbCrLf & _
                               r_str_NumDoc & " - " & r_str_NomPrv, vbExclamation, modgen_g_str_NomPlt
                        Exit Sub
                     End If
                     'fin mismo proveedor
                  End If
               End If
           Next
               
           For r_int_Fila = 0 To grd_Listad.Rows - 1
               If Trim(grd_Listad.TextMatrix(r_int_Fila, 12)) = "X" Then
                  .grd_Listad.Rows = .grd_Listad.Rows + 1
                  .grd_Listad.Row = .grd_Listad.Rows - 1
               
                  .grd_Listad.Col = 0 ''CODIGO
                  .grd_Listad.Text = grd_Listad.TextMatrix(r_int_Fila, 0) 'CODIGO
                  .grd_Listad.Col = 1 'TIPO PROCESO
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 1) & "")  'TIPO PROCESO
                  .grd_Listad.Col = 2 'FECHA - 1050
                  .grd_Listad.Text = grd_Listad.TextMatrix(r_int_Fila, 2) 'FECHA - 1050
                  .grd_Listad.Col = 3 'NRO DOCUMENTO
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 22) & "")  'NRO DOCUMENTO
                  .grd_Listad.Col = 4 'PROVEEDOR
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 4) & "")  'PROVEEDOR
                  .grd_Listad.Col = 5 'DESCRIPCION
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 5) & "")  'DESCRIPCION
                  
                  .grd_Listad.Col = 7 'MONEDA
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 7) & "") 'MONEDA
                  .grd_Listad.Col = 6 'CUENTA CORRIENTE
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 6) & "") 'CUENTA CORRIENTE
                  .grd_Listad.Col = 8 'IMPORTE
                  .grd_Listad.Text = grd_Listad.TextMatrix(r_int_Fila, 8) 'IMPORTE
                  
                  .grd_Listad.Col = 9 'APLICACION
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 9) & "") 'APLICACION
                  .grd_Listad.Col = 10 'DESCUENTO
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 10) & "") 'DESCUENTO
                  .grd_Listad.Col = 11 'PAGO NETO
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 11) & "") 'PAGO NETO
                  .grd_Listad.Col = 12 'CODIGO APLICACION
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 13) & "") 'CODIGO APLICACION
                  
                  .grd_Listad.Col = 13 'COMAUT_CODAUT
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 14) & "") 'COMAUT_CODAUT
                  .grd_Listad.Col = 14 'COMAUT_TIPDOC
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 15) & "") 'COMAUT_TIPDOC
                  .grd_Listad.Col = 15 'COMAUT_NUMDOC
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 16) & "") 'COMAUT_NUMDOC
                  .grd_Listad.Col = 16 'COMAUT_CODMON
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 17) & "") 'COMAUT_CODMON
                  .grd_Listad.Col = 17 'COMAUT_CODBNC
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 18) & "") 'COMAUT_CODBNC
                  .grd_Listad.Col = 18 'COMAUT_CTACTB
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 19) & "") 'COMAUT_CTACTB
                  .grd_Listad.Col = 19 'COMAUT_DATCTB
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 20) & "") 'COMAUT_DATCTB
                  .grd_Listad.Col = 20 'COMAUT_TIPOPE
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 21) & "") 'COMAUT_TIPOPE
                  
                  .grd_Listad.Col = 21 'MAEPRV_CODSIC
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 23) & "") 'COMAUT_CODSIC
                  .grd_Listad.Col = 22 'MAEPRV_TIPPER
                  .grd_Listad.Text = Trim(grd_Listad.TextMatrix(r_int_Fila, 24) & "") 'COMAUT_TIPPER
                  
               End If
           Next
        End If
       .grd_Listad.Redraw = True
       Call gs_UbiIniGrid(.grd_Listad)
   End With
   
   Unload Me
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call gs_LimpiaGrid(grd_Listad)
   
   grd_Listad.ColWidth(0) = 1080 'CODIGO
   grd_Listad.ColWidth(1) = 1370 'TIPO PROCESO
   grd_Listad.ColWidth(2) = 1030 'FECHA
   grd_Listad.ColWidth(3) = 1290 'USU_APRUEBA
   grd_Listad.ColWidth(4) = 2720 'PROVEEDOR
   grd_Listad.ColWidth(5) = 1670 'DESCRIPCION
   grd_Listad.ColWidth(6) = 1930 'CUENTA CORRIENTE
   grd_Listad.ColWidth(7) = 780  'MONEDA
   grd_Listad.ColWidth(8) = 1080 'TOTAL
   grd_Listad.ColWidth(9) = 930  '--APLICACION
   grd_Listad.ColWidth(10) = 780   '--DESCUENTO
   grd_Listad.ColWidth(11) = 1060 '--PAGO NETO
   grd_Listad.ColWidth(12) = 800 'SELECCIONAR
   grd_Listad.ColWidth(13) = 0 'APLICACION CODIGO
   grd_Listad.ColWidth(14) = 0 'COMAUT_CODAUT
   grd_Listad.ColWidth(15) = 0 'COMAUT_TIPDOC
   grd_Listad.ColWidth(16) = 0 'COMAUT_NUMDOC
   grd_Listad.ColWidth(17) = 0 'COMAUT_CODMON
   grd_Listad.ColWidth(18) = 0 'COMAUT_CODBNC
   grd_Listad.ColWidth(19) = 0 'COMAUT_CTACTB
   grd_Listad.ColWidth(20) = 0 'COMAUT_DATCTB
   grd_Listad.ColWidth(21) = 0 'COMAUT_TIPOPE
   grd_Listad.ColWidth(22) = 0 'NRO-DOCUMENTO
   grd_Listad.ColWidth(23) = 0 'CODIGO_PLANILLA
   grd_Listad.ColWidth(24) = 0 'TIPO_PERSONAL
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter 'CODIGO
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter 'PROVEEDOR
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignLeftCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter 'TOTAL
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
   grd_Listad.ColAlignment(11) = flexAlignRightCenter
   grd_Listad.ColAlignment(12) = flexAlignCenterCenter
   
   rb_Ninguno.Value = True
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
                  moddat_g_str_TipDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 15))
                  moddat_g_str_NumDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 16))
                  moddat_g_int_InsAct = 0 'tipo registro compra
                  frm_Ctb_RegCom_04.Show 1
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case "05" 'ENTREGAS A RENDIR
                  r_str_FlgAux = moddat_g_int_FlgGrb 'guardando estado
                  r_str_CodAux = moddat_g_str_Codigo 'guardando codigo
                  moddat_g_str_Codigo = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_str_CodIte = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
                  moddat_g_str_CodMod = grd_Listad.TextMatrix(grd_Listad.Row, 17)
                  moddat_g_int_FlgGrb = 0
                  If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 21) & "") = "1" Then
                     frm_Ctb_EntRen_02.Show 1 'form principal
                  ElseIf Trim(grd_Listad.TextMatrix(grd_Listad.Row, 21) & "") = "2" Then
                     frm_Ctb_EntRen_04.Show 1 'reembolso
                  End If
                  moddat_g_str_Codigo = r_str_CodAux 'devolviendo su origen
                  moddat_g_int_FlgGrb = r_str_FlgAux 'devolviendo su origen
          Case Else
               Exit Sub
   End Select
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_Buscar()
Dim r_str_Cadena  As String

   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)

   r_str_Cadena = ""
   For l_int_Contar = 0 To frm_Ctb_PagCom_02.grd_Listad.Rows - 1
       r_str_Cadena = r_str_Cadena & CLng(frm_Ctb_PagCom_02.grd_Listad.TextMatrix(l_int_Contar, 0)) & ","
   Next
   If r_str_Cadena <> "" Then
      r_str_Cadena = Mid(r_str_Cadena, 1, Len(r_str_Cadena) - 1)
   End If
  '--------------------------------------
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT COMAUT_CODAUT, COMAUT_CODOPE, COMAUT_FECOPE, COMAUT_TIPDOC, COMAUT_NUMDOC,  "
   g_str_Parame = g_str_Parame & "        COMAUT_CODMON, COMAUT_IMPPAG, COMAUT_CODBNC, COMAUT_CTACRR, COMAUT_CTACTB,  "
   g_str_Parame = g_str_Parame & "        COMAUT_DATCTB , COMAUT_CODEST, TRIM(C.PARDES_DESCRI) AS MONEDA,  "
   g_str_Parame = g_str_Parame & "        TRIM(D.PARDES_DESCRI) AS TIPOPROCESO, COMAUT_USUAPR, "
   g_str_Parame = g_str_Parame & "        DECODE(B.MaePrv_RazSoc,NULL,TRIM(E.DATGEN_APEPAT) ||' '|| TRIM(E.DATGEN_APEMAT) ||' '|| TRIM(E.DATGEN_NOMBRE) "
   g_str_Parame = g_str_Parame & "               ,B.MaePrv_RazSoc) AS MaePrv_RazSoc, A.COMAUT_TIPOPE, A.COMAUT_DESCRP,  "
   g_str_Parame = g_str_Parame & "        B.MAEPRV_CODSIC , B.MAEPRV_TIPPER "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_COMAUT A  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.COMAUT_TIPDOC AND TRIM(B.MAEPRV_NUMDOC) = TRIM(A.COMAUT_NUMDOC)  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = A.COMAUT_TIPDOC AND TRIM(E.DATGEN_NUMDOC) = TRIM(A.COMAUT_NUMDOC) "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.COMAUT_CODMON  "
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 136 AND TO_NUMBER(D.PARDES_CODITE) = TO_NUMBER(SUBSTR(LPAD(COMAUT_CODOPE,10,0),1,2)) AND D.PARDES_CODITE <> 0   "
   g_str_Parame = g_str_Parame & "  WHERE COMAUT_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "    AND COMAUT_CODEST = 2  "
   If r_str_Cadena <> "" Then
      g_str_Parame = g_str_Parame & "    AND COMAUT_CODOPE NOT IN (" & r_str_Cadena & ")  "
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY COMAUT_FECOPE ASC, A.COMAUT_CODAUT ASC  "
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
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!COMAUT_FECOPE)
   
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!COMAUT_USUAPR & "")
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!MaePrv_RazSoc & "")
                                    
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(CStr(g_rst_Princi!COMAUT_DESCRP & ""))
                  
      grd_Listad.Col = 6
      grd_Listad.Text = CStr(g_rst_Princi!COMAUT_CTACRR & "")
      
      grd_Listad.Col = 7
      grd_Listad.Text = CStr(g_rst_Princi!Moneda & "")
            
      grd_Listad.Col = 8 'TOTAL A PAGAR
      grd_Listad.Text = Format(g_rst_Princi!COMAUT_IMPPAG, "###,###,###,##0.00")
      
      grd_Listad.Col = 9
      grd_Listad.Text = "NINGUNO" 'APLICACION
      
      grd_Listad.Col = 10
      grd_Listad.Text = "0.00" 'DESCUENTO
      
      grd_Listad.Col = 11
      grd_Listad.Text = Format(g_rst_Princi!COMAUT_IMPPAG, "###,###,###,##0.00") 'PAGO NETO
                                 
      'grd_Listad.Col = 12
      'grd_Listad.Text = COLUMNA SELECCIONAR
                                 
      grd_Listad.Col = 13
      grd_Listad.Text = 1 'APLICACION CODIGO
      
      grd_Listad.Col = 14
      grd_Listad.Text = g_rst_Princi!COMAUT_CODAUT
      
      grd_Listad.Col = 15
      grd_Listad.Text = g_rst_Princi!COMAUT_TIPDOC
      
      grd_Listad.Col = 16
      grd_Listad.Text = g_rst_Princi!COMAUT_NUMDOC
      
      grd_Listad.Col = 17
      grd_Listad.Text = g_rst_Princi!COMAUT_CODMON
      
      grd_Listad.Col = 18
      grd_Listad.Text = Trim(CStr(g_rst_Princi!COMAUT_CODBNC & ""))
      
      grd_Listad.Col = 19
      grd_Listad.Text = g_rst_Princi!COMAUT_CTACTB
      
      grd_Listad.Col = 20
      grd_Listad.Text = g_rst_Princi!COMAUT_DATCTB
      
      grd_Listad.Col = 21
      grd_Listad.Text = g_rst_Princi!COMAUT_TIPOPE
      
      grd_Listad.Col = 22
      grd_Listad.Text = g_rst_Princi!COMAUT_TIPDOC & "-" & Trim(g_rst_Princi!COMAUT_NUMDOC & "")
      
      grd_Listad.Col = 23
      grd_Listad.Text = Trim(g_rst_Princi!MAEPRV_CODSIC & "")
      
      grd_Listad.Col = 24
      grd_Listad.Text = Trim(g_rst_Princi!MAEPRV_TIPPER & "")
      
      g_rst_Princi.MoveNext
   Loop

   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_Click()
   If grd_Listad.Rows = 0 Then
      rb_Ninguno.Value = True
      rb_ITF.Value = False
      rb_4TA.Value = False
      Exit Sub
   End If
   
   If grd_Listad.TextMatrix(grd_Listad.Row, 9) = "NINGUNO" Then
      rb_Ninguno.Value = True
      rb_ITF.Value = False
      rb_4TA.Value = False
   End If
   If grd_Listad.TextMatrix(grd_Listad.Row, 9) = "ITF" Then
      rb_Ninguno.Value = False
      rb_ITF.Value = True
      rb_4TA.Value = False
   End If
   If grd_Listad.TextMatrix(grd_Listad.Row, 9) = "4TA" Then
      rb_Ninguno.Value = False
      rb_ITF.Value = False
      rb_4TA.Value = True
   End If
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 12
   If grd_Listad.Text = "X" Then
       grd_Listad.Text = ""
   Else
        grd_Listad.Text = "X"
   End If
   Call gs_RefrescaGrid(grd_Listad)
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE PAGOS APROBADOS"
      .Range(.Cells(2, 2), .Cells(2, 14)).Merge
      .Range(.Cells(2, 2), .Cells(2, 14)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 14)).HorizontalAlignment = xlHAlignCenter

      .Cells(3, 2) = "CÓDIGO"
      .Cells(3, 3) = "TIPO PROCESO"
      .Cells(3, 4) = "FECHA"
      .Cells(3, 5) = "EVALUADOR"
      .Cells(3, 6) = "NRO DOCUMENTO"
      .Cells(3, 7) = "PROVEEDOR"
      .Cells(3, 8) = "DESCRIPCIÓN"
      .Cells(3, 9) = "CUENTA CORRIENTE"
      .Cells(3, 10) = "MONEDA"
      .Cells(3, 11) = "TOTAL PAGAR"
      .Cells(3, 12) = "APLICACION"
      .Cells(3, 13) = "DESCUENTO"
      .Cells(3, 14) = "PAGO NETO"
         
      .Range(.Cells(3, 2), .Cells(3, 14)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 14)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13 'codigo
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 22 'tipo proceso
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 13 'fecha
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 17 'EVALUDOR
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 17 'nro documento
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 45 'proveedor
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 22 'descripcion
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 22 'cuenta corriente
      .Columns("I").HorizontalAlignment = xlHAlignLeft
      .Columns("J").ColumnWidth = 22 'moneda
      .Columns("J").HorizontalAlignment = xlHAlignLeft
      .Columns("K").ColumnWidth = 14 'total a pagar
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 12 'APLICACION
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 12 'DESCUENTO
      .Columns("M").NumberFormat = "###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 14 'PAGO NETO
      .Columns("N").NumberFormat = "###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(10, 14)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 14)).Font.Size = 11
      
      r_int_NumFil = 4
      For l_int_Contar = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil, 2) = "'" & grd_Listad.TextMatrix(l_int_Contar, 0)
         .Cells(r_int_NumFil, 3) = grd_Listad.TextMatrix(l_int_Contar, 1)
         .Cells(r_int_NumFil, 4) = "'" & grd_Listad.TextMatrix(l_int_Contar, 2)
         .Cells(r_int_NumFil, 5) = grd_Listad.TextMatrix(l_int_Contar, 3)
         .Cells(r_int_NumFil, 6) = grd_Listad.TextMatrix(l_int_Contar, 22)
         .Cells(r_int_NumFil, 7) = grd_Listad.TextMatrix(l_int_Contar, 4)
         .Cells(r_int_NumFil, 8) = "'" & grd_Listad.TextMatrix(l_int_Contar, 5)
         .Cells(r_int_NumFil, 9) = "'" & grd_Listad.TextMatrix(l_int_Contar, 6)
         .Cells(r_int_NumFil, 10) = grd_Listad.TextMatrix(l_int_Contar, 7)
         .Cells(r_int_NumFil, 11) = grd_Listad.TextMatrix(l_int_Contar, 8)
         .Cells(r_int_NumFil, 12) = grd_Listad.TextMatrix(l_int_Contar, 9)
         .Cells(r_int_NumFil, 13) = grd_Listad.TextMatrix(l_int_Contar, 10)
         .Cells(r_int_NumFil, 14) = grd_Listad.TextMatrix(l_int_Contar, 11)
         r_int_NumFil = r_int_NumFil + 1
      Next
      .Range(.Cells(3, 3), .Cells(3, 14)).HorizontalAlignment = xlHAlignCenter
      
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub rb_4TA_Click()
Dim r_str_AplNom  As String
Dim r_int_AplCod  As Integer
Dim r_dbl_AplDsc  As String
Dim r_dbl_ImpNet  As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   Call frm_Ctb_PagCom_02.fs_Calc4TA(grd_Listad.TextMatrix(grd_Listad.Row, 8), r_str_AplNom, r_int_AplCod, r_dbl_AplDsc, r_dbl_ImpNet)
   
   grd_Listad.TextMatrix(grd_Listad.Row, 9) = r_str_AplNom
   grd_Listad.TextMatrix(grd_Listad.Row, 10) = r_dbl_AplDsc
   grd_Listad.TextMatrix(grd_Listad.Row, 11) = r_dbl_ImpNet
   grd_Listad.TextMatrix(grd_Listad.Row, 13) = r_int_AplCod
End Sub

Private Sub rb_ITF_Click()
Dim r_str_AplNom  As String
Dim r_int_AplCod  As Integer
Dim r_dbl_AplDsc  As String
Dim r_dbl_ImpNet  As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   Call frm_Ctb_PagCom_02.fs_CalcITF(grd_Listad.TextMatrix(grd_Listad.Row, 8), r_str_AplNom, r_int_AplCod, r_dbl_AplDsc, r_dbl_ImpNet)
   
   grd_Listad.TextMatrix(grd_Listad.Row, 9) = r_str_AplNom
   grd_Listad.TextMatrix(grd_Listad.Row, 10) = r_dbl_AplDsc
   grd_Listad.TextMatrix(grd_Listad.Row, 11) = r_dbl_ImpNet
   grd_Listad.TextMatrix(grd_Listad.Row, 13) = r_int_AplCod
End Sub

Private Sub rb_Ninguno_Click()
Dim r_str_AplNom  As String
Dim r_int_AplCod  As Integer
Dim r_dbl_AplDsc  As String
Dim r_dbl_ImpNet  As String

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   Call frm_Ctb_PagCom_02.fs_CalcNinguno(grd_Listad.TextMatrix(grd_Listad.Row, 8), r_str_AplNom, r_int_AplCod, r_dbl_AplDsc, r_dbl_ImpNet)
   
   grd_Listad.TextMatrix(grd_Listad.Row, 9) = r_str_AplNom
   grd_Listad.TextMatrix(grd_Listad.Row, 10) = r_dbl_AplDsc
   grd_Listad.TextMatrix(grd_Listad.Row, 11) = r_dbl_ImpNet
   grd_Listad.TextMatrix(grd_Listad.Row, 13) = r_int_AplCod 'CODIGO
End Sub
