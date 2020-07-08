VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_CajChc_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15360
   Icon            =   "GesCtb_frm_190.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8745
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   15555
      _Version        =   65536
      _ExtentX        =   27437
      _ExtentY        =   15425
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
         Height          =   5715
         Left            =   60
         TabIndex        =   8
         Top             =   2310
         Width           =   15255
         _Version        =   65536
         _ExtentX        =   26908
         _ExtentY        =   10081
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
         Begin Threed.SSPanel pnl_ImpTot 
            Height          =   285
            Left            =   12690
            TabIndex        =   36
            Top             =   5340
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   12582912
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
            Alignment       =   4
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   4965
            Left            =   30
            TabIndex        =   9
            Top             =   360
            Width           =   15195
            _ExtentX        =   26802
            _ExtentY        =   8758
            _Version        =   393216
            Rows            =   24
            Cols            =   11
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_DebMN 
            Height          =   285
            Left            =   8820
            TabIndex        =   10
            Top             =   60
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   1958
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Contable"
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
            Left            =   9900
            TabIndex        =   11
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Comprobante"
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
            Left            =   1080
            TabIndex        =   12
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Documento"
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
            Left            =   2490
            TabIndex        =   13
            Top             =   60
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Razón Social"
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
            TabIndex        =   14
            Top             =   60
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   12690
            TabIndex        =   15
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Total Comprobante"
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
            Left            =   11940
            TabIndex        =   16
            Top             =   60
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
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
            Left            =   14100
            TabIndex        =   35
            Top             =   60
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1402
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Procesado"
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
            Left            =   5940
            TabIndex        =   38
            Top             =   60
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Total ==>"
            Height          =   195
            Left            =   11850
            TabIndex        =   37
            Top             =   5370
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   17
         Top             =   60
         Width           =   15255
         _Version        =   65536
         _ExtentX        =   26908
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
            Height          =   495
            Left            =   630
            TabIndex        =   18
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Detalle de Caja Chica"
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
            Picture         =   "GesCtb_frm_190.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   19
         Top             =   780
         Width           =   15255
         _Version        =   65536
         _ExtentX        =   26908
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
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_190.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_190.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1260
            Picture         =   "GesCtb_frm_190.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14640
            Picture         =   "GesCtb_frm_190.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_190.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   2460
            Picture         =   "GesCtb_frm_190.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   825
         Left            =   60
         TabIndex        =   21
         Top             =   1470
         Width           =   15255
         _Version        =   65536
         _ExtentX        =   26908
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
         Begin Threed.SSPanel pnl_NumCaja 
            Height          =   315
            Left            =   1230
            TabIndex        =   30
            Top             =   90
            Width           =   2355
            _Version        =   65536
            _ExtentX        =   4154
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   1230
            TabIndex        =   31
            Top             =   420
            Width           =   2355
            _Version        =   65536
            _ExtentX        =   4154
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
         Begin Threed.SSPanel pnl_Importe 
            Height          =   315
            Left            =   5070
            TabIndex        =   32
            Top             =   420
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Respon 
            Height          =   315
            Left            =   7830
            TabIndex        =   33
            Top             =   420
            Width           =   5775
            _Version        =   65536
            _ExtentX        =   10186
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
         Begin Threed.SSPanel pnl_FechaCaja 
            Height          =   315
            Left            =   5070
            TabIndex        =   34
            Top             =   90
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
         Begin Threed.SSPanel pnl_Asigna 
            Height          =   315
            Left            =   7830
            TabIndex        =   39
            Top             =   90
            Width           =   5775
            _Version        =   65536
            _ExtentX        =   10186
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
         Begin VB.Label lbl_Asigna 
            AutoSize        =   -1  'True
            Caption         =   "Asignación:"
            Height          =   195
            Left            =   6780
            TabIndex        =   40
            Top             =   180
            Width           =   825
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Responsable:"
            Height          =   195
            Left            =   6780
            TabIndex        =   29
            Top             =   510
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mto Asignado:"
            Height          =   195
            Left            =   3930
            TabIndex        =   28
            Top             =   510
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nro de Caja:"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   180
            Width           =   885
         End
         Begin VB.Label lbl_Fecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Caja:"
            Height          =   195
            Left            =   3930
            TabIndex        =   23
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   510
            Width           =   630
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   570
         Left            =   60
         TabIndex        =   24
         Top             =   8070
         Width           =   15255
         _Version        =   65536
         _ExtentX        =   26908
         _ExtentY        =   1005
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
            Left            =   5400
            MaxLength       =   100
            TabIndex        =   6
            Top             =   150
            Width           =   4425
         End
         Begin VB.ComboBox cmb_Buscar 
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   150
            Width           =   2595
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Columna a Buscar:"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   26
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Por:"
            Height          =   195
            Left            =   4530
            TabIndex        =   25
            Top             =   210
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_CajChc_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type arr_CajDet
   CajDet_CodDet        As String
   CajDet_CodCaj        As String
   CajDet_FecEmi        As String
   CajDet_TipCpb    As Long
   CajDet_TipCpb_Lrg    As String
   CajDet_Nserie        As String
   CajDet_NroCom        As String
   CajDet_TipDoc_Lrg    As String
   CajDet_NumDoc        As String
   MaePrv_RazSoc        As String
   CajDet_Moneda        As String
'   CajDet_Grv1          As Double
'   CajDet_Grv2          As Double
   'CajDet_Ngv           As Double
'   CajDet_Igv           As Double
   CajDet_Ppg           As Double
   CajChc_FecCaj        As String
   CAJDET_FECCTB        As String
End Type
   
Dim l_arr_CajDet()      As arr_CajDet

Private Sub cmb_Buscar_Click()
   If (cmb_Buscar.ListIndex = 0 Or cmb_Buscar.ListIndex = -1) Then
       txt_Buscar.Enabled = False
       Call fs_BuscarCaja
   Else
       txt_Buscar.Enabled = True
       Call gs_SetFocus(txt_Buscar)
   End If
   txt_Buscar.Text = ""
End Sub

Private Sub cmb_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_Buscar.Enabled = False) Then
          Call fs_BuscarCaja
      Else
          Call gs_SetFocus(txt_Buscar)
      End If
   End If
End Sub

Private Sub cmd_Borrar_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_CodIte = "" 'CAJDET_CODDET
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 8
   If UCase(Trim(grd_Listad.Text)) = "SI" Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "No se pudo eliminar el registro por que esta procesado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If (moddat_g_int_Situac = 1) Then
       MsgBox "No se pudo eliminar el registro, se encuentra procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
   If MsgBox("¿Seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_CodIte = grd_Listad.Text
   Call gs_RefrescaGrid(grd_Listad)
   
   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC_DET_BORRAR ( "
   g_str_Parame = g_str_Parame & "'" & Trim(moddat_g_str_CodIte) & "', " 'CAJDET_CODDET
   g_str_Parame = g_str_Parame & "'" & Trim(moddat_g_str_Codigo) & "', " 'CAJDET_CODCAJ
   
   'tipo de proceso
   If moddat_g_int_TipEva = 1 Then
      g_str_Parame = g_str_Parame & " 1, "
   ElseIf moddat_g_int_TipEva = 6 Then
      g_str_Parame = g_str_Parame & " 6, "
   End If
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo completar la eliminación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   Else
      MsgBox "El registro se elimino correctamente.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
   
   Call fs_BuscarCaja
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_ExpExc_Click()
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_BuscarCaja
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "NINGUNA"
   cmb_Buscar.AddItem "RAZÓN SOCIAL"
   cmb_Buscar.AddItem "PROCESADO"
   
   If moddat_g_int_TipEva = 1 Then
      'caja chica
      pnl_NumCaja.Caption = Format(moddat_g_str_Codigo, "0000000000") 'nro caja
      pnl_FechaCaja.Caption = moddat_g_str_FecIng 'fecha caja
      pnl_Respon.Caption = moddat_g_str_Descri 'responsable
      pnl_Moneda.Caption = moddat_g_str_DesMod 'moneda
      pnl_Importe.Caption = Format(moddat_g_dbl_MtoPre, "###,###,###,##0.00") & " " 'importe
      lbl_Asigna.Enabled = False
      pnl_Asigna.Enabled = False
      lbl_Fecha.Caption = "Fecha Caja:"
      pnl_Titulo.Caption = "Detalle de Caja Chica"
   ElseIf moddat_g_int_TipEva = 6 Then
      'tarjeta credito
      pnl_NumCaja.Caption = Format(moddat_g_str_Codigo, "0000000000") 'nro caja
      pnl_FechaCaja.Caption = moddat_g_str_FecIng 'fecha caja
      pnl_Respon.Caption = moddat_g_str_Descri 'responsable
      pnl_Asigna.Caption = moddat_g_str_NomPrd 'asigna
      pnl_Moneda.Caption = moddat_g_str_DesMod 'moneda
      pnl_Importe.Caption = Format(moddat_g_dbl_MtoPre, "###,###,###,##0.00") & " " 'importe
      lbl_Fecha.Caption = "Periodo:"
      pnl_Titulo.Caption = "Detalle de Tarjeta de Crédito"
   End If
   
   grd_Listad.ColWidth(0) = 1030 'codigo
   grd_Listad.ColWidth(1) = 1410 'ID-CLIENTE
   grd_Listad.ColWidth(2) = 3450 'razon Social
   grd_Listad.ColWidth(3) = 2870 'Descripcion
   grd_Listad.ColWidth(4) = 1100 'fecha contable
   grd_Listad.ColWidth(5) = 2020 'tipo comprobante
   grd_Listad.ColWidth(6) = 750 'moneda
   grd_Listad.ColWidth(7) = 1400 'total compronte
   grd_Listad.ColWidth(8) = 800 'procesado
   grd_Listad.ColWidth(9) = 0 'tipo documento
   grd_Listad.ColWidth(10) = 0 'nro documento
      
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   cmb_Buscar.ListIndex = 0
   pnl_ImpTot.Caption = "0.00"
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_Agrega_Click()
   If (moddat_g_int_Situac = 1) Then
       Call gs_RefrescaGrid(grd_Listad)
       MsgBox "No se pudo adicionar un registro, se encuentra procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   
   'caja chica
   moddat_g_int_FlgGrb = 1 'insert
   frm_Ctb_CajChc_04.Show 1
End Sub

Private Sub cmd_Consul_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_CodIte = "" 'CAJDET_CODDET
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   Call gs_RefrescaGrid(grd_Listad)
   
   grd_Listad.Col = 0
   moddat_g_str_CodIte = CStr(grd_Listad.Text)
   grd_Listad.Col = 9
   moddat_g_str_TipDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 10
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)

   moddat_g_int_FlgGrb = 0 'consultar
   frm_Ctb_CajChc_04.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_CodIte = "" 'CAJDET_CODDET
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If (moddat_g_int_Situac = 1) Then
       Call gs_RefrescaGrid(grd_Listad)
       MsgBox "No se pudo editar el registro, se encuentra procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   
   grd_Listad.Col = 0 'CAJDET_CODDET
   moddat_g_str_CodIte = CStr(grd_Listad.Text)
   grd_Listad.Col = 9
   moddat_g_str_TipDoc = CStr(grd_Listad.Text)
   grd_Listad.Col = 10
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   
   moddat_g_int_FlgGrb = 2 'editar
   Call gs_RefrescaGrid(grd_Listad)
   frm_Ctb_CajChc_04.Show 1
   
   Call fs_BuscarCaja
   Call gs_SetFocus(grd_Listad)
End Sub

Public Sub fs_BuscarCaja()
Dim r_str_Cadena  As String
Dim r_dbl_Import  As Double
Dim r_str_TipTab  As String

   ReDim l_arr_CajDet(0)
   pnl_ImpTot.Caption = "0.00"
   r_dbl_Import = 0
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   
   r_str_TipTab = 1
   If moddat_g_int_TipEva = 1 Then
      r_str_TipTab = "1"
   ElseIf moddat_g_int_TipEva = 6 Then
      r_str_TipTab = "6"
   End If

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.CAJDET_CODDET, A.CAJDET_TIPDOC || '-' || A.CAJDET_NUMDOC ID_CLIENTE, TRIM(CAJDET_DESCRP) CAJDET_DESCRP,  "
   g_str_Parame = g_str_Parame & "       B.CAJCHC_FECCAJ, A.CAJDET_FLGPRC,DECODE(A.CAJDET_FLGPRC,1,'SI','NO') AS PROCESADO,  "
   g_str_Parame = g_str_Parame & "       A.CAJDET_CODCAJ, A.CAJDET_FECEMI, TRIM(D.PARDES_DESCRI) TIPO_COMPROBANTE, CAJDET_TIPCPB, "
   g_str_Parame = g_str_Parame & "       A.CAJDET_NSERIE, A.CAJDET_NROCOM, TRIM(F.PARDES_DESCRI) TIPO_DOCUMENTO, A.CAJDET_TIPDOC,  "
   g_str_Parame = g_str_Parame & "       A.CAJDET_NUMDOC, TRIM(C.MAEPRV_RAZSOC) MAEPRV_RAZSOC, TRIM(E.PARDES_DESCRI) MONEDA, "
   'g_str_Parame = g_str_Parame & "       CAJDET_DEB_GRV1, CAJDET_HAB_GRV1, CAJDET_DEB_GRV2, CAJDET_HAB_GRV2,  "
   'g_str_Parame = g_str_Parame & "       CAJDET_DEB_NGV1, CAJDET_HAB_NGV1, CAJDET_DEB_IGV1, CAJDET_HAB_IGV1,  "
   'g_str_Parame = g_str_Parame & "       CAJDET_DEB_PPG1 , CAJDET_HAB_PPG1, CAJDET_FECCTB  "
   g_str_Parame = g_str_Parame & "       (CASE WHEN B.CAJCHC_CODMON = A.CAJDET_CODMON THEN (NVL(CAJDET_DEB_PPG1,0) + NVL(CAJDET_HAB_PPG1,0)) "
   g_str_Parame = g_str_Parame & "             WHEN A.CAJDET_CODMON = 1 THEN (NVL(CAJDET_DEB_PPG1,0) + NVL(CAJDET_HAB_PPG1,0)) / G.TIPCAM_VENTAS "
   g_str_Parame = g_str_Parame & "             WHEN A.CAJDET_CODMON = 2 THEN (NVL(CAJDET_DEB_PPG1,0) + NVL(CAJDET_HAB_PPG1,0)) * G.TIPCAM_VENTAS "
   g_str_Parame = g_str_Parame & "         END) AS CAJDET_TOTPPG, G.TIPCAM_VENTAS AS TIPCAM_SBS, CAJDET_FECCTB "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_CAJCHC_DET A  "
   g_str_Parame = g_str_Parame & " INNER JOIN CNTBL_CAJCHC B ON A.CAJDET_CODCAJ = B.CAJCHC_CODCAJ AND A.CAJDET_TIPTAB = " & r_str_TipTab & " AND B.CAJCHC_TIPTAB = " & r_str_TipTab
   g_str_Parame = g_str_Parame & " INNER JOIN CNTBL_MAEPRV C ON A.CAJDET_TIPDOC = C.MAEPRV_TIPDOC AND A.CAJDET_NUMDOC = C.MAEPRV_NUMDOC "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 123 AND A.CAJDET_TIPCPB = D.PARDES_CODITE " 'comprobante
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 204 AND A.CAJDET_CODMON = E.PARDES_CODITE " 'moneda
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = 118 AND A.CAJDET_TIPDOC = F.PARDES_CODITE " 'documento
   g_str_Parame = g_str_Parame & "  LEFT JOIN OPE_TIPCAM G ON TIPCAM_CODIGO = 3 AND TIPCAM_TIPMON = 2 AND G.TIPCAM_FECDIA = A.CAJDET_FECEMI "
   g_str_Parame = g_str_Parame & " WHERE A.CAJDET_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND A.CAJDET_TIPTAB = " & r_str_TipTab
   g_str_Parame = g_str_Parame & "   AND A.CAJDET_CODCAJ = " & moddat_g_str_Codigo
      
   If (cmb_Buscar.ListIndex = 1) Then 'razon social
       If Len(Trim(txt_Buscar.Text)) > 0 Then
           g_str_Parame = g_str_Parame & "   AND UPPER(TRIM(C.MAEPRV_RAZSOC)) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'"
       End If
   ElseIf (cmb_Buscar.ListIndex = 2) Then 'procesado
       r_str_Cadena = ""
       Select Case UCase(Trim(txt_Buscar.Text))
              Case "S", "SI", "I": r_str_Cadena = "1"
              Case "N", "NO", "O": r_str_Cadena = "0"
       End Select
       If (Len(Trim(r_str_Cadena)) > 0) Then
           g_str_Parame = g_str_Parame & "   AND CAJDET_FLGPRC = " & r_str_Cadena
       End If
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY CAJDET_CODDET ASC  "
   
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
   ReDim l_arr_CajDet(0)

   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1

      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!CajDet_CodDet)

      grd_Listad.Col = 1
      grd_Listad.Text = CStr(g_rst_Princi!ID_CLIENTE & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!MaePrv_RazSoc & "")
            
      grd_Listad.Col = 3
      grd_Listad.Text = CStr(g_rst_Princi!CAJDET_DESCRP & "")
            
      grd_Listad.Col = 4
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!CAJDET_FECCTB) 'CajChc_FecCaj)

      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!TIPO_COMPROBANTE & "")
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(moddat_g_str_DesMod)  'moneda del registro principal
      'grd_Listad.Text = Trim(g_rst_Princi!Moneda & "")
            
      grd_Listad.Col = 7
      grd_Listad.Text = Format(g_rst_Princi!CajDet_TotPpg, "###,###,###,##0.00")
      'grd_Listad.Text = Format(g_rst_Princi!CAJDET_Deb_Ppg1 + g_rst_Princi!CAJDET_Hab_Ppg1, "###,###,###,##0.00")
      
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(g_rst_Princi!PROCESADO & "")
      
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(g_rst_Princi!CAJDET_TipDoc & "")
      
      grd_Listad.Col = 10
      grd_Listad.Text = Trim(g_rst_Princi!CajDet_NumDoc & "")

      '***AGREGAR AL ARREGLO
      ReDim Preserve l_arr_CajDet(UBound(l_arr_CajDet) + 1)
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_CodDet = g_rst_Princi!CajDet_CodDet
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_CodCaj = g_rst_Princi!CajDet_CodCaj
      l_arr_CajDet(UBound(l_arr_CajDet)).CajChc_FecCaj = g_rst_Princi!CajChc_FecCaj
      l_arr_CajDet(UBound(l_arr_CajDet)).CAJDET_FECCTB = g_rst_Princi!CAJDET_FECCTB
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_FecEmi = g_rst_Princi!CajDet_FecEmi
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_NumDoc = g_rst_Princi!CajDet_NumDoc
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TipCpb_Lrg = g_rst_Princi!TIPO_COMPROBANTE
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Nserie = g_rst_Princi!CajDet_Nserie
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_NroCom = g_rst_Princi!CajDet_NroCom
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TipDoc_Lrg = g_rst_Princi!TIPO_DOCUMENTO
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_NumDoc = g_rst_Princi!CajDet_NumDoc
      l_arr_CajDet(UBound(l_arr_CajDet)).MaePrv_RazSoc = g_rst_Princi!MaePrv_RazSoc
      'l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Moneda = g_rst_Princi!Moneda
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Moneda = Trim(moddat_g_str_DesMod) 'moneda del registro principal
      'l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Grv1 = g_rst_Princi!CAJDET_Deb_Grv1 + g_rst_Princi!CAJDET_Hab_Grv1
      'l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Grv2 = g_rst_Princi!CAJDET_Deb_Grv2 + g_rst_Princi!CAJDET_Hab_Grv2
      'l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Ngv = g_rst_Princi!CAJDET_Deb_Ngv1 + g_rst_Princi!CAJDET_Hab_Ngv1
      'l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Igv = g_rst_Princi!CAJDET_Deb_Igv1 + g_rst_Princi!CAJDET_Hab_Igv1
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Ppg = g_rst_Princi!CajDet_TotPpg
      
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TipCpb = g_rst_Princi!CajDet_TipCpb
      If g_rst_Princi!CajDet_TipCpb = 7 Or g_rst_Princi!CajDet_TipCpb = 88 Then
         r_dbl_Import = r_dbl_Import - (g_rst_Princi!CajDet_TotPpg)
         'r_dbl_Import = r_dbl_Import - (g_rst_Princi!CAJDET_Deb_Ppg1 + g_rst_Princi!CAJDET_Hab_Ppg1)
      Else
         r_dbl_Import = r_dbl_Import + (g_rst_Princi!CajDet_TotPpg)
         'r_dbl_Import = r_dbl_Import + (g_rst_Princi!CAJDET_Deb_Ppg1 + g_rst_Princi!CAJDET_Hab_Ppg1)
      End If
      g_rst_Princi.MoveNext
   Loop

   pnl_ImpTot.Caption = Format(r_dbl_Import, "###,###,###,##0.00")
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_dbl_MtoImp        As Double
Dim r_int_Contar        As Integer

   r_dbl_MtoImp = 0
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      If moddat_g_int_TipEva = 1 Then 'tipo de tabla
         .Cells(1, 2) = "Caja " & Format(moddat_g_str_Codigo, "00000000")
         .Cells(2, 2) = "REPORTE DE CAJA CHICA"
      ElseIf moddat_g_int_TipEva = 6 Then 'tipo de tabla
         .Cells(1, 2) = "Tarjeta " & Format(moddat_g_str_Codigo, "00000000")
         .Cells(2, 2) = "REPORTE DE TARJETA DE CREDITO CORPORATIVO"
      End If
      
      .Cells(1, 11) = "Fecha: " & moddat_g_str_FecIng & " "
      .Range(.Cells(2, 2), .Cells(2, 11)).Merge
      .Range(.Cells(1, 2), .Cells(2, 11)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 11)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(3, 2), .Cells(3, 4)).Merge
      .Range(.Cells(4, 2), .Cells(4, 4)).Merge
      .Range(.Cells(6, 2), .Cells(6, 4)).Merge
      .Range(.Cells(2, 6), .Cells(2, 8)).Font.Bold = True
      .Cells(3, 2) = "Responsable: " & moddat_g_str_Descri
      .Cells(4, 2) = "Moneda: " & moddat_g_str_DesMod
      .Cells(6, 2) = "Detalle del consumo"
            
      .Cells(7, 2) = "CÓDIGO"
      .Cells(7, 3) = "FECHA EMISIÓN"
      .Cells(7, 4) = "TIPO COMPROBANTE"
      .Cells(7, 5) = "SERIE"
      .Cells(7, 6) = "NÚMERO"
      .Cells(7, 7) = "DOCUMENTO"
      .Range(.Cells(7, 8), .Cells(7, 10)).Merge
      .Cells(7, 8) = "PROVEEDOR"
      .Cells(7, 11) = "TOTAL COMPROBANTE"
      .Range(.Cells(7, 8), .Cells(7, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 2), .Cells(7, 11)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(6, 2), .Cells(7, 11)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 7 'codigo
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12 'fecha de emision
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 18 'tipo de comprobante
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 6 'serie
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9 'numero
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 13 'documento
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 14 'proveedor
      .Columns("I").ColumnWidth = 7 'proveedor
      .Columns("J").ColumnWidth = 13 'proveedor
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("K").ColumnWidth = 20 'total
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(1, 1), .Cells(11, 11)).Font.Name = "Calibri"
      
      r_int_NumFil = 6
      r_dbl_MtoImp = 0
      For r_int_Contar = 1 To UBound(l_arr_CajDet)
          .Cells(r_int_NumFil + 2, 2) = "'" & l_arr_CajDet(r_int_Contar).CajDet_CodDet 'codigo
          .Cells(r_int_NumFil + 2, 3) = "'" & gf_FormatoFecha(l_arr_CajDet(r_int_Contar).CajDet_FecEmi) 'fecha de emision
          .Cells(r_int_NumFil + 2, 4) = "'" & l_arr_CajDet(r_int_Contar).CajDet_TipCpb_Lrg 'tipo de comprobante
          .Cells(r_int_NumFil + 2, 5) = "'" & l_arr_CajDet(r_int_Contar).CajDet_Nserie     'serie
          .Cells(r_int_NumFil + 2, 6) = "'" & l_arr_CajDet(r_int_Contar).CajDet_NroCom     'numero
          .Cells(r_int_NumFil + 2, 7) = "'" & l_arr_CajDet(r_int_Contar).CajDet_NumDoc     'documento
          .Range(.Cells(r_int_NumFil + 2, 8), .Cells(r_int_NumFil + 2, 10)).Merge
          .Cells(r_int_NumFil + 2, 8) = "'" & l_arr_CajDet(r_int_Contar).MaePrv_RazSoc     'proveedor
          .Cells(r_int_NumFil + 2, 11) = l_arr_CajDet(r_int_Contar).CajDet_Ppg
          
          If l_arr_CajDet(r_int_Contar).CajDet_TipCpb = 7 Or l_arr_CajDet(r_int_Contar).CajDet_TipCpb = 88 Then
             r_dbl_MtoImp = r_dbl_MtoImp - l_arr_CajDet(r_int_Contar).CajDet_Ppg
          Else
            r_dbl_MtoImp = r_dbl_MtoImp + l_arr_CajDet(r_int_Contar).CajDet_Ppg
          End If
          
          r_int_NumFil = r_int_NumFil + 1
      Next
      .Cells(r_int_NumFil + 2, 10) = "Por reembolsar "
      .Cells(r_int_NumFil + 4, 10) = "Asignado "
      .Cells(r_int_NumFil + 6, 10) = "Saldo en Caja "
      .Cells(r_int_NumFil + 2, 11) = Format(r_dbl_MtoImp, "###,###,###,##0.00")
      .Cells(r_int_NumFil + 4, 11) = Format(moddat_g_dbl_MtoPre, "###,###,###,##0.00")
      .Cells(r_int_NumFil + 6, 11) = Format(moddat_g_dbl_MtoPre - r_dbl_MtoImp, "###,###,###,##0.00")
      .Range(.Cells(r_int_NumFil + 2, 10), .Cells(r_int_NumFil + 6, 11)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(r_int_NumFil + 2, 10), .Cells(r_int_NumFil + 6, 11)).Font.Bold = True
      
      .Range(.Cells(1, 1), .Cells(r_int_NumFil + 6, 11)).Font.Size = 10
      .Range(.Cells(r_int_NumFil + 12, 2), .Cells(r_int_NumFil + 12, 4)).Merge
      .Range(.Cells(r_int_NumFil + 12, 6), .Cells(r_int_NumFil + 12, 8)).Merge
      .Range(.Cells(r_int_NumFil + 12, 10), .Cells(r_int_NumFil + 12, 11)).Merge
      .Cells(r_int_NumFil + 12, 2) = "Gerente de Administración y finanzas"
      .Cells(r_int_NumFil + 12, 6) = "Responsable"
      .Cells(r_int_NumFil + 12, 10) = "Contabilidad"
            
      .Range(.Cells(r_int_NumFil + 12, 2), .Cells(r_int_NumFil + 12, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 12, 6), .Cells(r_int_NumFil + 12, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 12, 10), .Cells(r_int_NumFil + 12, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 12, 2), .Cells(r_int_NumFil + 12, 11)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 2)).HorizontalAlignment = xlHAlignLeft
      
      .Cells(r_int_NumFil + 14, 11) = "Fecha de Reporte: " & moddat_g_str_FecSis
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub txt_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call fs_BuscarCaja
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub
