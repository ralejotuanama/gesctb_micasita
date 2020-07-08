VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_EntRen_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   Icon            =   "GesCtb_frm_194.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15345
      _Version        =   65536
      _ExtentX        =   27067
      _ExtentY        =   16060
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
         Height          =   3600
         Left            =   60
         TabIndex        =   1
         Top             =   2310
         Width           =   15135
         _Version        =   65536
         _ExtentX        =   26696
         _ExtentY        =   6350
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
            Left            =   10410
            TabIndex        =   34
            Top             =   3240
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2558
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
            Height          =   2625
            Left            =   70
            TabIndex        =   2
            Top             =   570
            Width           =   15060
            _ExtentX        =   26564
            _ExtentY        =   4630
            _Version        =   393216
            Rows            =   24
            Cols            =   15
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_FecCon 
            Height          =   285
            Left            =   6030
            TabIndex        =   3
            Top             =   300
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
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
         Begin Threed.SSPanel pnl_TipCpb 
            Height          =   285
            Left            =   7050
            TabIndex        =   4
            Top             =   300
            Width           =   2530
            _Version        =   65536
            _ExtentX        =   4463
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
         Begin Threed.SSPanel pnl_NumDoc 
            Height          =   285
            Left            =   1230
            TabIndex        =   5
            Top             =   300
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
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
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   285
            Left            =   2550
            TabIndex        =   6
            Top             =   300
            Width           =   3495
            _Version        =   65536
            _ExtentX        =   6174
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   285
            Left            =   90
            TabIndex        =   7
            Top             =   300
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
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
         Begin Threed.SSPanel pnl_TotCom 
            Height          =   285
            Left            =   10460
            TabIndex        =   8
            Top             =   300
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
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
         Begin Threed.SSPanel pnl_Proces 
            Height          =   285
            Left            =   11850
            TabIndex        =   10
            Top             =   300
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
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
         Begin Threed.SSPanel pnl_FecPag 
            Height          =   285
            Left            =   12720
            TabIndex        =   36
            Top             =   300
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Pago"
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
         Begin Threed.SSPanel pnl_CodPag 
            Height          =   285
            Left            =   13755
            TabIndex        =   37
            Top             =   300
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Pago"
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
         Begin Threed.SSPanel pnl_TipMon 
            Height          =   285
            Left            =   9570
            TabIndex        =   9
            Top             =   300
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
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
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Relación Documento Sustentados"
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
            Left            =   90
            TabIndex        =   43
            Top             =   60
            Width           =   2895
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Total ==>"
            Height          =   195
            Left            =   9690
            TabIndex        =   35
            Top             =   3270
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   15130
         _Version        =   65536
         _ExtentX        =   26688
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
            TabIndex        =   12
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Detalle de Entregas a Rendir"
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
            Picture         =   "GesCtb_frm_194.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   13
         Top             =   770
         Width           =   15135
         _Version        =   65536
         _ExtentX        =   26688
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
         Begin VB.CommandButton cmb_Reembolso 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_194.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Reembolso"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmb_Devolucion 
            Height          =   585
            Left            =   2490
            Picture         =   "GesCtb_frm_194.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Devolución"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   3720
            Picture         =   "GesCtb_frm_194.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   3120
            Picture         =   "GesCtb_frm_194.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Consultar Comprobante"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14520
            Picture         =   "GesCtb_frm_194.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1260
            Picture         =   "GesCtb_frm_194.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Eliminar Comprobante"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_194.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Modificar Comprobante"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_194.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Adicionar Comprobante"
            Top             =   30
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   825
         Left            =   60
         TabIndex        =   20
         Top             =   1440
         Width           =   15130
         _Version        =   65536
         _ExtentX        =   26688
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
            TabIndex        =   21
            Top             =   90
            Width           =   2355
            _Version        =   65536
            _ExtentX        =   4154
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   1230
            TabIndex        =   22
            Top             =   420
            Width           =   2355
            _Version        =   65536
            _ExtentX        =   4154
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
         Begin Threed.SSPanel pnl_Importe 
            Height          =   315
            Left            =   5070
            TabIndex        =   23
            Top             =   420
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Respon 
            Height          =   315
            Left            =   7830
            TabIndex        =   24
            Top             =   420
            Width           =   5775
            _Version        =   65536
            _ExtentX        =   10186
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
         Begin Threed.SSPanel pnl_FechaCaja 
            Height          =   315
            Left            =   5070
            TabIndex        =   25
            Top             =   90
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   510
            Width           =   630
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de ER:"
            Height          =   195
            Left            =   3930
            TabIndex        =   29
            Top             =   180
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nro de ER:"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mto Asignado:"
            Height          =   195
            Left            =   3930
            TabIndex        =   27
            Top             =   510
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Responsable:"
            Height          =   195
            Left            =   6780
            TabIndex        =   26
            Top             =   510
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   60
         TabIndex        =   31
         Top             =   5955
         Width           =   15135
         _Version        =   65536
         _ExtentX        =   26696
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
         Begin VB.CommandButton cmd_Consul_TipPag 
            Height          =   585
            Left            =   1260
            Picture         =   "GesCtb_frm_194.frx":1C9E
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Consultar Tipo de Pago"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar_TipPag 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_194.frx":1FA8
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Eliminar Tipo de Pago"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega_TipPag 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_194.frx":22B2
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Adicionar Tipo de Pago"
            Top             =   30
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2400
         Left            =   60
         TabIndex        =   39
         Top             =   6645
         Width           =   15135
         _Version        =   65536
         _ExtentX        =   26696
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
         BevelOuter      =   1
         Begin Threed.SSPanel pnl_TotCab 
            Height          =   285
            Left            =   10410
            TabIndex        =   40
            Top             =   2070
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2558
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
         Begin MSFlexGridLib.MSFlexGrid grd_PagAsc 
            Height          =   1470
            Left            =   60
            TabIndex        =   41
            Top             =   570
            Width           =   15060
            _ExtentX        =   26564
            _ExtentY        =   2593
            _Version        =   393216
            Rows            =   24
            Cols            =   12
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   90
            TabIndex        =   45
            Top             =   270
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro ER"
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
            Left            =   3345
            TabIndex        =   46
            Top             =   270
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Responsable"
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
            Left            =   9495
            TabIndex        =   47
            Top             =   270
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
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
         Begin Threed.SSPanel pnl_MtoAsig 
            Height          =   285
            Left            =   10380
            TabIndex        =   48
            Top             =   270
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto Asignado"
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
         Begin Threed.SSPanel pnl_FecEnt 
            Height          =   285
            Left            =   1230
            TabIndex        =   49
            Top             =   270
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha de ER"
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
         Begin Threed.SSPanel pnl_Glosa 
            Height          =   285
            Left            =   7695
            TabIndex        =   50
            Top             =   270
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Glosa"
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
         Begin Threed.SSPanel pnl_Benefi 
            Height          =   285
            Left            =   5520
            TabIndex        =   51
            Top             =   270
            Width           =   2190
            _Version        =   65536
            _ExtentX        =   3863
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Benficiario"
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
         Begin Threed.SSPanel pnl_TipPag 
            Height          =   285
            Left            =   2250
            TabIndex        =   52
            Top             =   270
            Width           =   1110
            _Version        =   65536
            _ExtentX        =   1958
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Pago"
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
            Left            =   11820
            TabIndex        =   53
            Top             =   270
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
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
            Left            =   12690
            TabIndex        =   54
            Top             =   270
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Pago"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   13725
            TabIndex        =   55
            Top             =   270
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Código Pago"
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipos de Pagos Asociados"
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
            Left            =   60
            TabIndex        =   44
            Top             =   30
            Width           =   2265
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total ==>"
            Height          =   195
            Left            =   9690
            TabIndex        =   42
            Top             =   2100
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_EntRen_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type arr_CajDet
   CajDet_CodDet        As String
   CajDet_CodCaj        As String
   CajDet_FecEmi        As String
   CajDet_TipCpb        As String
   CajDet_TipCpb_Lrg    As String
   CajDet_Nserie        As String
   CajDet_NroCom        As String
   CajDet_TipDoc_Lrg    As String
   CajDet_NumDoc        As String
   MaePrv_RazSoc        As String
   CajDet_Moneda        As String
   CajDet_TotPpg        As Double
   CajChc_FecCaj        As String
End Type
   
Dim l_arr_CajDet()      As arr_CajDet

Private Sub cmb_Devolucion_Click()
   If (moddat_g_int_Situac = 1) Then
       MsgBox "El registro de entrega a rendir esta procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If

'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT CAJCHC_TIPTAB  "
'   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC  "
'   g_str_Parame = g_str_Parame & "  WHERE CajChc_CodCaj = " & CLng(pnl_NumCaja.Caption)
'   g_str_Parame = g_str_Parame & "    AND CAJCHC_TIPTAB IN (4,5)  "
'   g_str_Parame = g_str_Parame & "    AND CAJCHC_SITUAC = 1  "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'      Screen.MousePointer = 0
'      Exit Sub
'   End If
'
'   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'ningún registro
'      g_rst_Princi.Close
'      Set g_rst_Princi = Nothing
'      Screen.MousePointer = 0
'
'      moddat_g_int_FlgGrb = 1 'adicionar
'      frm_Ctb_EntRen_05.Show 1
'      Exit Sub
'   End If

   'g_rst_Princi.MoveFirst
   'If g_rst_Princi!CAJCHC_TIPTAB = 4 Then
   '   MsgBox "El registro de devolución ya fue ingresado.", vbExclamation, modgen_g_str_NomPlt
   '   Exit Sub
   'End If
   'If g_rst_Princi!CAJCHC_TIPTAB = 5 Then
   '   MsgBox "Un registro de reembolso ya fue ingresado.", vbExclamation, modgen_g_str_NomPlt
   '   Exit Sub
   'End If
   
   moddat_g_int_FlgGrb = 1 'adicionar
   frm_Ctb_EntRen_05.Show 1
End Sub

Private Sub cmb_Reembolso_Click()
   If (moddat_g_int_Situac = 1) Then
       MsgBox "El registro de entrega a rendir esta procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If

'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & " SELECT CAJCHC_TIPTAB  "
'   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC  "
'   g_str_Parame = g_str_Parame & "  WHERE CajChc_CodCaj = " & CLng(pnl_NumCaja.Caption)
'   g_str_Parame = g_str_Parame & "    AND CAJCHC_TIPTAB IN (4,5)  "
'   g_str_Parame = g_str_Parame & "    AND CAJCHC_SITUAC = 1  "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'      Screen.MousePointer = 0
'      Exit Sub
'   End If
'
'   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'ningún registro
'      g_rst_Princi.Close
'      Set g_rst_Princi = Nothing
'      Screen.MousePointer = 0
'
'      moddat_g_int_FlgGrb = 1 'adicionar
'      frm_Ctb_EntRen_04.Show 1
'      Exit Sub
'   End If
'
'   g_rst_Princi.MoveFirst
   'If g_rst_Princi!CAJCHC_TIPTAB = 4 Then
   '   MsgBox "El registro de devolución ya fue ingresado.", vbExclamation, modgen_g_str_NomPlt
   '   Exit Sub
   'End If
   'If g_rst_Princi!CAJCHC_TIPTAB = 5 Then
   '   MsgBox "Un registro de reembolso ya fue ingresado.", vbExclamation, modgen_g_str_NomPlt
   '   Exit Sub
   'End If
   
   moddat_g_int_FlgGrb = 1 'adicionar
   frm_Ctb_EntRen_04.Show 1
End Sub

Private Sub cmd_Agrega_TipPag_Click()
   If (moddat_g_int_Situac = 1) Then
       Call gs_RefrescaGrid(grd_PagAsc)
       MsgBox "El registro de entrega a rendir esta procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   
   frm_Ctb_EntRen_06.Show 1
End Sub

Private Sub cmd_Borrar_Click()
Dim r_bol_Estado    As Boolean

   r_bol_Estado = False
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_CodIte = "" 'CAJDET_CODDET
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If UCase(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 7))) = "SI" Then
      MsgBox "El registro de entrega a rendir esta procesado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If (moddat_g_int_Situac = 1) Then
       MsgBox "El registro de entrega a rendir esta procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   
   'CAJDET_CODDET
   moddat_g_str_CodIte = grd_Listad.TextMatrix(grd_Listad.Row, 0)
   
   If grd_Listad.TextMatrix(grd_Listad.Row, 12) <> 1 Then
      If frm_Ctb_EntRen_01.fs_ValMod_Aut(moddat_g_str_CodIte, 2, CStr(grd_Listad.TextMatrix(grd_Listad.Row, 13))) = False Then
         Exit Sub
      End If
   End If
   
   If grd_Listad.TextMatrix(grd_Listad.Row, 12) = 1 Then
      'facturas u otros
      If MsgBox("¿Seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      Screen.MousePointer = 11
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC_DET_BORRAR ( "
      g_str_Parame = g_str_Parame & "'" & Trim(moddat_g_str_CodIte) & "', " 'CAJDET_CODDET
      g_str_Parame = g_str_Parame & "'" & CLng(moddat_g_str_Codigo) & "', " 'CAJDET_CODCAJ
      g_str_Parame = g_str_Parame & " 2, "
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
         r_bol_Estado = True
      End If
      Screen.MousePointer = 0
   Else
      If grd_Listad.TextMatrix(grd_Listad.Row, 12) = 4 Then
         If MsgBox("¿Seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         End If
       
         'devolucion
         Screen.MousePointer = 11
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC_BORRAR ( "
         g_str_Parame = g_str_Parame & "'" & Trim(moddat_g_str_Codigo) & "', " 'CAJCHC_CODCAJ
         g_str_Parame = g_str_Parame & "4, " 'CAJCHC_TIPTAB
         g_str_Parame = g_str_Parame & grd_Listad.TextMatrix(grd_Listad.Row, 13) & ", " 'CAJCHC_NUMERO
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
            r_bol_Estado = True
         End If
         Screen.MousePointer = 0
      End If
      If grd_Listad.TextMatrix(grd_Listad.Row, 12) = 5 Then
         'REEMBOLSO(1)
         If MsgBox("¿Seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         End If
         Screen.MousePointer = 11
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC_BORRAR ( "
         g_str_Parame = g_str_Parame & "'" & Trim(moddat_g_str_Codigo) & "', " 'CAJCHC_CODCAJ
         g_str_Parame = g_str_Parame & "5, " 'CAJCHC_TIPTAB
         g_str_Parame = g_str_Parame & grd_Listad.TextMatrix(grd_Listad.Row, 13) & ", " 'CAJCHC_NUMERO
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
            r_bol_Estado = True
         End If
         Screen.MousePointer = 0
      End If
   End If
   
   If r_bol_Estado = True Then
      Call fs_BuscarCaja
      Call frm_Ctb_EntRen_01.fs_BuscarCaja
      Call gs_SetFocus(grd_Listad)
   End If
End Sub

Private Sub cmd_Borrar_TipPag_Click()
   If grd_PagAsc.Rows = 0 Then
      Exit Sub
   End If
   
   If (moddat_g_int_Situac = 1) Then
       Call gs_RefrescaGrid(grd_PagAsc)
       MsgBox "El registro de entrega a rendir esta procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   
   If CLng(grd_PagAsc.TextMatrix(grd_PagAsc.Row, 11)) = 1 Then 'Registro Origen
      MsgBox "Este registro no se puede eiminar.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Call gs_RefrescaGrid(grd_PagAsc)
   If MsgBox("¿Seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If CLng(grd_PagAsc.TextMatrix(grd_PagAsc.Row, 11)) = 2 Then
      'grd_PagAsc.RemoveItem (grd_PagAsc.Row)
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC_ASOC ( "
      g_str_Parame = g_str_Parame & CLng(moddat_g_str_Codigo) & ", "
      g_str_Parame = g_str_Parame & CLng(grd_PagAsc.TextMatrix(grd_PagAsc.Row, 0)) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(2) & ") "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
   
      MsgBox "El registro se elimino correctamente.", vbInformation, modgen_g_str_NomPlt
      Call frm_Ctb_EntRen_01.fs_BuscarCaja
      Call fs_BuscarAsoc
   End If
End Sub

Private Sub cmd_Consul_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_CodIte = "" 'CAJDET_CODDET
   moddat_g_dbl_IngDec = 0 ' ITEM
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_str_CodIte = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   moddat_g_str_TipDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 8))
   moddat_g_str_NumDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 9))

   moddat_g_int_FlgGrb = 0 'consultar
   If CInt(grd_Listad.TextMatrix(grd_Listad.Row, 12)) = 1 Then
      moddat_g_int_InsAct = 1 'entregas a rendir
      frm_Ctb_RegCom_04.Show 1
   ElseIf CInt(grd_Listad.TextMatrix(grd_Listad.Row, 12)) = 4 Then
      moddat_g_dbl_IngDec = grd_Listad.TextMatrix(grd_Listad.Row, 13)
      frm_Ctb_EntRen_05.Show 1 'devolucion
   ElseIf CInt(grd_Listad.TextMatrix(grd_Listad.Row, 12)) = 5 Then
      moddat_g_dbl_IngDec = grd_Listad.TextMatrix(grd_Listad.Row, 13)
      frm_Ctb_EntRen_04.Show 1 'reembolso
   End If
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Consul_TipPag_Click()
Dim r_str_Codigo As String
Dim r_int_FlgGrb As Integer
   
   If grd_PagAsc.Rows = 0 Then
      Exit Sub
   End If
   
   r_str_Codigo = moddat_g_str_Codigo
   r_int_FlgGrb = moddat_g_int_FlgGrb

   moddat_g_str_Codigo = CLng(grd_PagAsc.TextMatrix(grd_PagAsc.Row, 0))
   moddat_g_int_FlgGrb = 0 'consultar
   frm_Ctb_EntRen_02.Show 1
   moddat_g_str_Codigo = r_str_Codigo
   moddat_g_int_FlgGrb = r_int_FlgGrb
   
   Call gs_SetFocus(grd_PagAsc)
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   moddat_g_str_CodIte = "" 'CAJDET_CODDET
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If moddat_g_int_Situac = 1 Then
      Call gs_RefrescaGrid(grd_Listad)
      MsgBox "El registro de entrega a rendir esta procesado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'SOLO SE EDITAN ENTREGAS A RENDIR DETALLE
   If CInt(grd_Listad.TextMatrix(grd_Listad.Row, 12)) <> 1 Then
      Exit Sub
   End If
   
   moddat_g_str_CodIte = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 0)) 'CAJDET_CODDET
   moddat_g_str_TipDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 8))
   moddat_g_str_NumDoc = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 9))
      
   moddat_g_int_FlgGrb = 2 'editar
   moddat_g_int_InsAct = 1 'entregas a rendir
   frm_Ctb_RegCom_04.Show 1
         
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call fs_BuscarCaja
   Call fs_BuscarAsoc
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   
   pnl_NumCaja.Caption = moddat_g_str_Codigo
   pnl_FechaCaja.Caption = moddat_g_str_FecIng
   pnl_Moneda.Caption = moddat_g_str_DesMod
   pnl_Importe.Caption = Format(moddat_g_dbl_MtoPre, "###,###,###,##0.00") & " "
   pnl_Respon.Caption = moddat_g_str_Descri
   
   With frm_Ctb_EntRen_01
        'SI ES REEMBOLSO
        If Trim(CStr(.grd_Listad.TextMatrix(.grd_Listad.Row, 18) & "")) <> "1" Then
           cmd_Agrega_TipPag.Enabled = False
           cmd_Borrar_TipPag.Enabled = False
        End If
        'SI FUE PROCESADO
        If (moddat_g_int_Situac = 1) Then
           cmd_Agrega_TipPag.Enabled = False
           cmd_Borrar_TipPag.Enabled = False
           cmd_Agrega.Enabled = False
           cmd_Editar.Enabled = False
           cmd_Borrar.Enabled = False
           cmb_Devolucion.Enabled = False
           cmb_Reembolso.Enabled = False
        End If
        If Trim(CStr(.grd_Listad.TextMatrix(.grd_Listad.Row, 18) & "")) = "1" Then
           'SI ES ANTICIPO
           If Trim(.grd_Listad.TextMatrix(.grd_Listad.Row, 13)) = "" And CInt(.grd_Listad.TextMatrix(.grd_Listad.Row, 23)) > 0 Then
              cmd_Agrega.Enabled = False
              cmd_Editar.Enabled = False
              cmb_Devolucion.Enabled = False
              cmb_Reembolso.Enabled = False
              cmd_Agrega_TipPag.Enabled = False
           End If
        End If
   End With
   
   grd_Listad.ColWidth(0) = 1140 'codigo
   grd_Listad.ColWidth(1) = 1340 'nro Documento
   grd_Listad.ColWidth(2) = 3470 'razon Social
   grd_Listad.ColWidth(3) = 1030 'fecha contable
   grd_Listad.ColWidth(4) = 2520 'tipo comprobante
   grd_Listad.ColWidth(5) = 870 'moneda
   grd_Listad.ColWidth(6) = 1400 'total compronte
   grd_Listad.ColWidth(7) = 880 'procesado
   grd_Listad.ColWidth(8) = 0 'tipo_doc
   grd_Listad.ColWidth(9) = 0 'num_doc¿
   grd_Listad.ColWidth(10) = 1020 'fecha pago
   grd_Listad.ColWidth(11) = 1040 'codigo pago
   grd_Listad.ColWidth(12) = 0 'tipo form - 1=entegas a rendir, 2=devolucion,3=reembolso
   grd_Listad.ColWidth(13) = 0 'ITEM-CORRELATIVO(devolucion reembolso)
   grd_Listad.ColWidth(14) = 0 'Asiento Contable
      
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignCenterCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
   
   grd_PagAsc.ColWidth(0) = 1160 'Nro ER
   grd_PagAsc.ColWidth(1) = 1020 'FECHA ER
   grd_PagAsc.ColWidth(2) = 1080 'TIPO PAGO
   grd_PagAsc.ColWidth(3) = 2190 'RESPONSABLE
   grd_PagAsc.ColWidth(4) = 2160 'BENEFICIRIO
   grd_PagAsc.ColWidth(5) = 1800 'GLOSA
   grd_PagAsc.ColWidth(6) = 890 'MONEDA
   grd_PagAsc.ColWidth(7) = 1440 'MTO ASIGNADO
   grd_PagAsc.ColWidth(8) = 890 'PROCESADO
   grd_PagAsc.ColWidth(9) = 1030 'FECHA PAGO
   grd_PagAsc.ColWidth(10) = 1030 'CODIGO PAGO
   grd_PagAsc.ColWidth(11) = 0 'Registro Cabecera
   
   grd_PagAsc.ColAlignment(0) = flexAlignCenterCenter
   grd_PagAsc.ColAlignment(1) = flexAlignCenterCenter
   grd_PagAsc.ColAlignment(2) = flexAlignLeftCenter
   grd_PagAsc.ColAlignment(3) = flexAlignLeftCenter
   grd_PagAsc.ColAlignment(4) = flexAlignLeftCenter
   grd_PagAsc.ColAlignment(5) = flexAlignLeftCenter 'GLOSA
   grd_PagAsc.ColAlignment(6) = flexAlignLeftCenter 'MONEDA
   grd_PagAsc.ColAlignment(7) = flexAlignRightCenter
   grd_PagAsc.ColAlignment(8) = flexAlignCenterCenter
   grd_PagAsc.ColAlignment(9) = flexAlignCenterCenter
   grd_PagAsc.ColAlignment(10) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   pnl_ImpTot.Caption = "0.00" & " "
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_LimpiaGrid(grd_PagAsc)
End Sub

Private Sub cmd_Agrega_Click()
   If (moddat_g_int_Situac = 1) Then
       Call gs_RefrescaGrid(grd_Listad)
       MsgBox "El registro de entrega a rendir esta procesado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   
   moddat_g_int_FlgGrb = 1 'insert
   moddat_g_int_InsAct = 1 'entregas a rendir
   frm_Ctb_RegCom_04.Show 1
End Sub

Public Sub fs_BuscarCaja()
Dim r_str_Cadena  As String
Dim r_dbl_Import  As Double

   ReDim l_arr_CajDet(0)
   pnl_ImpTot.Caption = "0.00" & " "
   r_dbl_Import = 0
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.CAJDET_CODDET, A.CAJDET_TIPDOC || '-' || A.CAJDET_NUMDOC ID_CLIENTE,  "
   g_str_Parame = g_str_Parame & "          B.CAJCHC_FECCAJ, DECODE(A.CAJDET_FLGPRC,1,'SI','NO') AS PROCESADO,  "
   g_str_Parame = g_str_Parame & "          A.CAJDET_CODCAJ, A.CAJDET_FECEMI, TRIM(D.PARDES_DESCRI) TIPO_COMPROBANTE, CAJDET_TIPCPB,  "
   g_str_Parame = g_str_Parame & "          A.CAJDET_NSERIE, A.CAJDET_NROCOM, TRIM(F.PARDES_DESCRI) TIPO_DOCUMENTO, A.CAJDET_TIPDOC,  "
   g_str_Parame = g_str_Parame & "          A.CAJDET_NUMDOC, TRIM(C.MAEPRV_RAZSOC) MAEPRV_RAZSOC, TRIM(E.PARDES_DESCRI) MONEDA,  "
   g_str_Parame = g_str_Parame & "          (CASE WHEN B.CAJCHC_CODMON = A.CAJDET_CODMON THEN (NVL(CAJDET_DEB_PPG1,0) + NVL(CAJDET_HAB_PPG1,0))  "
   g_str_Parame = g_str_Parame & "               WHEN A.CAJDET_CODMON = 1 THEN (NVL(CAJDET_DEB_PPG1,0) + NVL(CAJDET_HAB_PPG1,0)) / G.TIPCAM_VENTAS  "
   g_str_Parame = g_str_Parame & "               WHEN A.CAJDET_CODMON = 2 THEN (NVL(CAJDET_DEB_PPG1,0) + NVL(CAJDET_HAB_PPG1,0)) * G.TIPCAM_VENTAS  "
   g_str_Parame = g_str_Parame & "            END) AS CAJDET_TOTPPG, G.TIPCAM_VENTAS AS TIPCAM_SBS  "
   g_str_Parame = g_str_Parame & "     FROM CNTBL_CAJCHC_DET A  "
   g_str_Parame = g_str_Parame & "    INNER JOIN CNTBL_CAJCHC B ON A.CAJDET_CODCAJ = B.CAJCHC_CODCAJ AND A.CAJDET_TIPTAB = 2 AND B.CAJCHC_TIPTAB = 2  "
   g_str_Parame = g_str_Parame & "    INNER JOIN CNTBL_MAEPRV C ON A.CAJDET_TIPDOC = C.MAEPRV_TIPDOC AND A.CAJDET_NUMDOC = C.MAEPRV_NUMDOC  "
   g_str_Parame = g_str_Parame & "    INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 123 AND A.CAJDET_TIPCPB = D.PARDES_CODITE  " 'comprobante
   g_str_Parame = g_str_Parame & "    INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 204 AND A.CAJDET_CODMON = E.PARDES_CODITE  " 'moneda
   g_str_Parame = g_str_Parame & "    INNER JOIN MNT_PARDES F ON F.PARDES_CODGRP = 118 AND A.CAJDET_TIPDOC = F.PARDES_CODITE  " 'documento
   'g_str_Parame = g_str_Parame & "     LEFT JOIN OPE_TIPCAM G ON TIPCAM_CODIGO = 2 AND TIPCAM_TIPMON = 2 AND G.TIPCAM_FECDIA = A.CAJDET_FECEMI  "
   g_str_Parame = g_str_Parame & "     LEFT JOIN OPE_TIPCAM G ON TIPCAM_CODIGO = 3 AND TIPCAM_TIPMON = 2 AND G.TIPCAM_FECDIA = A.CAJDET_FECEMI  "
   g_str_Parame = g_str_Parame & "    WHERE A.CAJDET_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "      AND A.CAJDET_TIPTAB = 2  "
   g_str_Parame = g_str_Parame & "      AND A.CAJDET_CODCAJ = " & CLng(moddat_g_str_Codigo)
   g_str_Parame = g_str_Parame & "  ORDER BY CAJDET_CODDET ASC  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call fs_Devol_remb(r_dbl_Import)
      pnl_ImpTot.Caption = Format(r_dbl_Import, "###,###,###,##0.00") & " "
      If grd_Listad.Rows > 0 Then
         grd_Listad.Redraw = True
         Call gs_UbiIniGrid(grd_Listad)
      End If
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
      grd_Listad.Text = Format(g_rst_Princi!CajDet_CodDet, "0000000000")

      grd_Listad.Col = 1
      grd_Listad.Text = CStr(g_rst_Princi!ID_CLIENTE & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!MaePrv_RazSoc & "")
            
      grd_Listad.Col = 3
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!CajChc_FecCaj)

      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!TIPO_COMPROBANTE & "")
      
      'moneda del registro principal
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(moddat_g_str_DesMod) 'Trim(g_rst_Princi!Moneda & "")
            
      grd_Listad.Col = 6
      grd_Listad.Text = Format(g_rst_Princi!CajDet_TotPpg, "###,###,###,##0.00")
      
      grd_Listad.Col = 7
      grd_Listad.Text = Trim(g_rst_Princi!PROCESADO & "")
      
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(g_rst_Princi!CAJDET_TipDoc & "")
      
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(g_rst_Princi!CajDet_NumDoc & "")
      
      'grd_Listad.Col = 10
      'grd_Listad.Text = 'fecha pago
      
      'grd_Listad.Col = 11
      'grd_Listad.Text = 'codigo pago
      
      grd_Listad.Col = 12
      grd_Listad.Text = 1 'tipo form - 1=entegas a rendir, 2=devolucion,3=reembolso

      '***AGREGAR AL ARREGLO
      ReDim Preserve l_arr_CajDet(UBound(l_arr_CajDet) + 1)
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_CodDet = g_rst_Princi!CajDet_CodDet
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_CodCaj = g_rst_Princi!CajDet_CodCaj
      l_arr_CajDet(UBound(l_arr_CajDet)).CajChc_FecCaj = g_rst_Princi!CajChc_FecCaj
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_FecEmi = g_rst_Princi!CajDet_FecEmi
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_NumDoc = g_rst_Princi!CajDet_NumDoc
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TipCpb_Lrg = g_rst_Princi!TIPO_COMPROBANTE
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TipCpb = g_rst_Princi!CajDet_TipCpb
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Nserie = g_rst_Princi!CajDet_Nserie
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_NroCom = g_rst_Princi!CajDet_NroCom
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TipDoc_Lrg = g_rst_Princi!TIPO_DOCUMENTO
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_NumDoc = g_rst_Princi!CajDet_NumDoc
      l_arr_CajDet(UBound(l_arr_CajDet)).MaePrv_RazSoc = g_rst_Princi!MaePrv_RazSoc
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Moneda = g_rst_Princi!Moneda
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TotPpg = g_rst_Princi!CajDet_TotPpg
   
      If g_rst_Princi!CajDet_TipCpb = 7 Or g_rst_Princi!CajDet_TipCpb = 88 Then
         r_dbl_Import = r_dbl_Import - (g_rst_Princi!CajDet_TotPpg)
      Else
         r_dbl_Import = r_dbl_Import + (g_rst_Princi!CajDet_TotPpg)
      End If
      
      g_rst_Princi.MoveNext
   Loop
         
   'REEMBOLSO Y DEVOLUCION SE AGREGAN AL FINAL
   Call fs_Devol_remb(r_dbl_Import)

   pnl_ImpTot.Caption = Format(r_dbl_Import, "###,###,###,##0.00") & " "
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Public Sub fs_BuscarAsoc()
'Dim r_int_fila  As Integer
Dim r_dbl_SumTot  As Double

   Call gs_LimpiaGrid(grd_PagAsc)
'   r_int_fila = 0
'
'   With frm_Ctb_EntRen_01
'      r_int_fila = .grd_Listad.Row
'      grd_PagAsc.Rows = grd_PagAsc.Rows + 1
'      grd_PagAsc.Row = grd_PagAsc.Rows - 1
'
'      grd_PagAsc.Col = 0 'CODIGO
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 0) 'CODIGO
'
'      grd_PagAsc.Col = 1 'fecha
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 1) 'fecha
'
'      grd_PagAsc.Col = 2 'tipo pago
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 2) 'tipo pago
'
'      grd_PagAsc.Col = 3 'responsable
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 3) 'responsable
'
'      grd_PagAsc.Col = 4 'beneficiario
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 4) 'beneficiario
'
'      grd_PagAsc.Col = 5 'glosa
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 5) 'glosa
'
'      grd_PagAsc.Col = 6 'moneda
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 6) 'moneda
'
'      grd_PagAsc.Col = 7 'monto asignado
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 7) 'monto asignado
'
'      grd_PagAsc.Col = 8 'proceso
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 10) 'proceso
'
'      grd_PagAsc.Col = 9 'fecha pago - compensacion
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 12) 'fecha pago - compensacion
'
'      grd_PagAsc.Col = 10 'codigo pago - compensacion
'      grd_PagAsc.Text = .grd_Listad.TextMatrix(r_int_fila, 13) 'codigo pago - compensacion
'
'      grd_PagAsc.Col = 11 'registro cabecera
'      grd_PagAsc.Text = 1 '1= cabecera, 2=Otros
'
'      Call gs_UbiIniGrid(grd_PagAsc)
'   End With
   
   r_dbl_SumTot = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.CAJCHC_CODCAJ, A.CAJCHC_FECCAJ, A.CAJCHC_TIPDOC || '-' || A.CAJCHC_NUMDOC NUMDOC_RESPON, TRIM(B.MAEPRV_RAZSOC) NOM_RESPON,  "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_CODMON, TRIM(C.PARDES_DESCRI) MONEDA, A.CAJCHC_IMPORT, A.CAJCHC_FLGPRC,  "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_DESCRI, A.CAJCHC_NUMOPE,  "
   g_str_Parame = g_str_Parame & "        TRIM(D.MAEPRV_RAZSOC) AS NOM_BENEFI, TRIM(F.PARDES_DESCRI) AS TIPO_PAGO, A.CAJCHC_TIPPAG,  "
   g_str_Parame = g_str_Parame & "        j.COMPAG_FECPAG , j.COMPAG_CODCOM, 1 AS REG_CAB  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.CAJCHC_TIPDOC AND B.MAEPRV_NUMDOC = A.CAJCHC_NUMDOC  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.CAJCHC_CODMON  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV D ON D.MAEPRV_TIPDOC = A.CAJCHC_TIPDOC_2 AND D.MAEPRV_NUMDOC = A.CAJCHC_NUMDOC_2  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 138 AND F.PARDES_CODITE = A.CAJCHC_TIPPAG  "  'tipo pago
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMAUT H ON TO_NUMBER(H.COMAUT_CODOPE) = TO_NUMBER(A.CAJCHC_CODCAJ) AND H.COMAUT_TIPOPE = 1 AND H.COMAUT_CODEST NOT IN (3)  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMDET I ON I.COMDET_CODAUT = H.COMAUT_CODAUT AND I.COMDET_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMPAG J ON J.COMPAG_CODCOM = I.COMDET_CODCOM AND J.COMPAG_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "                                                                 AND J.COMPAG_FLGCTB = 1  "
   g_str_Parame = g_str_Parame & "  WHERE A.CAJCHC_TIPTAB = 2  " 'entregas a rendir
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_CODCAJ = " & CLng(moddat_g_str_Codigo)
   g_str_Parame = g_str_Parame & "  UNION  "
   g_str_Parame = g_str_Parame & " SELECT A.CAJCHC_CODCAJ, A.CAJCHC_FECCAJ, A.CAJCHC_TIPDOC || '-' || A.CAJCHC_NUMDOC NUMDOC_RESPON, TRIM(B.MAEPRV_RAZSOC) NOM_RESPON,  "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_CODMON, TRIM(C.PARDES_DESCRI) MONEDA, A.CAJCHC_IMPORT, A.CAJCHC_FLGPRC,  "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_DESCRI, A.CAJCHC_NUMOPE,  "
   g_str_Parame = g_str_Parame & "        TRIM(D.MAEPRV_RAZSOC) AS NOM_BENEFI, TRIM(F.PARDES_DESCRI) AS TIPO_PAGO, A.CAJCHC_TIPPAG,  "
   g_str_Parame = g_str_Parame & "        j.COMPAG_FECPAG , j.COMPAG_CODCOM, 2 AS REG_CAB  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.CAJCHC_TIPDOC AND B.MAEPRV_NUMDOC = A.CAJCHC_NUMDOC  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.CAJCHC_CODMON  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV D ON D.MAEPRV_TIPDOC = A.CAJCHC_TIPDOC_2 AND D.MAEPRV_NUMDOC = A.CAJCHC_NUMDOC_2  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 138 AND F.PARDES_CODITE = A.CAJCHC_TIPPAG  "  'tipo pago
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMAUT H ON TO_NUMBER(H.COMAUT_CODOPE) = TO_NUMBER(A.CAJCHC_CODCAJ) AND H.COMAUT_TIPOPE = 1 AND H.COMAUT_CODEST NOT IN (3)  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMDET I ON I.COMDET_CODAUT = H.COMAUT_CODAUT AND I.COMDET_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMPAG J ON J.COMPAG_CODCOM = I.COMDET_CODCOM AND J.COMPAG_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "                                                                 AND J.COMPAG_FLGCTB = 1  "
   g_str_Parame = g_str_Parame & "  WHERE A.CAJCHC_TIPTAB = 2  " 'entregas a rendir
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_CODREF_1 = " & CLng(moddat_g_str_Codigo)
   g_str_Parame = g_str_Parame & "  ORDER BY REG_CAB ASC  "
   
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
   grd_PagAsc.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_PagAsc.Rows = grd_PagAsc.Rows + 1
      grd_PagAsc.Row = grd_PagAsc.Rows - 1
   
      grd_PagAsc.Col = 0
      grd_PagAsc.Text = Format(CStr(g_rst_Princi!CajChc_CodCaj), "0000000000")

      grd_PagAsc.Col = 1
      grd_PagAsc.Text = gf_FormatoFecha(g_rst_Princi!CajChc_FecCaj)

      grd_PagAsc.Col = 2
      grd_PagAsc.Text = CStr(g_rst_Princi!TIPO_PAGO & "")

      grd_PagAsc.Col = 3
      grd_PagAsc.Text = CStr(g_rst_Princi!NOM_RESPON & "")
      
      grd_PagAsc.Col = 4
      grd_PagAsc.Text = CStr(g_rst_Princi!NOM_BENEFI & "")
                  
      grd_PagAsc.Col = 5
      grd_PagAsc.Text = CStr(g_rst_Princi!CajChc_Descri & "")
      
      grd_PagAsc.Col = 6
      grd_PagAsc.Text = Trim(g_rst_Princi!Moneda & "")
      
      grd_PagAsc.Col = 7 'MTO ASIGNADO
      grd_PagAsc.Text = Format(g_rst_Princi!CajChc_Import, "###,###,###,##0.00")
      r_dbl_SumTot = r_dbl_SumTot + CDbl(g_rst_Princi!CajChc_Import)
      '--------------------------------------------------------------------------------------------------
      grd_PagAsc.Col = 8
      grd_PagAsc.Text = IIf(g_rst_Princi!CAJCHC_FLGPRC = 1, "SI", "NO")
            
      If Trim(g_rst_Princi!COMPAG_FECPAG & "") <> "" Then
         grd_PagAsc.Col = 9
         grd_PagAsc.Text = gf_FormatoFecha(g_rst_Princi!COMPAG_FECPAG)
      End If
      If Trim(g_rst_Princi!COMPAG_CODCOM & "") <> "" Then
         grd_PagAsc.Col = 10
         grd_PagAsc.Text = Format(g_rst_Princi!COMPAG_CODCOM, "00000000")
      End If
      
      grd_PagAsc.Col = 11 'registro cabecera
      grd_PagAsc.Text = g_rst_Princi!REG_CAB '1= cabecera, 2=Otros
      
      g_rst_Princi.MoveNext
   Loop
               
   moddat_g_dbl_MtoPre = r_dbl_SumTot 'importe
   pnl_Importe.Caption = Format(r_dbl_SumTot, "###,###,###,##0.00") & " "
   pnl_TotCab = Format(r_dbl_SumTot, "###,###,###,##0.00") & " "
   
   Call gs_UbiIniGrid(grd_PagAsc)
   grd_PagAsc.Redraw = True
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub fs_Devol_remb(ByRef p_Importe As Double)
Dim r_str_FecPag As String
Dim r_str_CodPag As String

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.CAJCHC_CODCAJ, A.CAJCHC_FECCAJ, A.CAJCHC_IMPORT, A.CAJCHC_DATCNT,  "
   g_str_Parame = g_str_Parame & "        CAJCHC_TIPCAM, CAJCHC_TIPDOC_2, CAJCHC_NUMDOC_2, CAJCHC_CODBCO_2, CAJCHC_CTACRR_2,  "
   g_str_Parame = g_str_Parame & "        TRIM(B.MAEPRV_RAZSOC) MAEPRV_RAZSOC,  "
   g_str_Parame = g_str_Parame & "        TRIM(C.PARDES_DESCRI) TIPO_DOCUMENTO, CAJCHC_TIPTAB, DECODE(A.CAJCHC_FLGPRC,1,'SI','NO') AS PROCESADO,  "
   g_str_Parame = g_str_Parame & "        (CASE WHEN A.CAJCHC_TIPTAB = 4 THEN 'DEVOLUCION'  " '--devolucion
   g_str_Parame = g_str_Parame & "              WHEN A.CAJCHC_TIPTAB = 5 THEN 'REEMBOLSO'  " '--reembolso
   g_str_Parame = g_str_Parame & "         END) AS NOM_DEVREM,  "
   g_str_Parame = g_str_Parame & "         j.COMPAG_FECPAG , j.COMPAG_CODCOM, A.CAJCHC_NUMERO  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC A  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.CAJCHC_TIPDOC_2 AND B.MAEPRV_NUMDOC = A.CAJCHC_NUMDOC_2  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES C ON C.PARDES_CODGRP = 118 AND C.PARDES_CODITE = A.CAJCHC_TIPDOC_2  " 'documento
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMAUT H ON TO_NUMBER(H.COMAUT_CODOPE) = TO_NUMBER(A.CAJCHC_CODCAJ) AND H.COMAUT_TIPOPE = 2 AND H.COMAUT_CODEST NOT IN (3)  "
   'g_str_Parame = g_str_Parame & "        AND TRIM(H.COMAUT_DATCTB) = TRIM(A.CAJCHC_DATCNT)  " 'checar
   g_str_Parame = g_str_Parame & "    AND TRIM(TO_CHAR(H.COMAUT_ORIGEN)) = TRIM(TO_CHAR(A.CAJCHC_NUMERO))  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMDET I ON I.COMDET_CODAUT = H.COMAUT_CODAUT AND I.COMDET_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMPAG J ON J.COMPAG_CODCOM = I.COMDET_CODCOM AND J.COMPAG_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "                                                                 AND J.COMPAG_FLGCTB = 1  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES K ON K.PARDES_CODGRP = 135 AND K.PARDES_CODITE = J.COMPAG_TIPPAG  "
   g_str_Parame = g_str_Parame & "  WHERE A.CajChc_CodCaj =  " & moddat_g_str_Codigo
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_TIPTAB IN (4,5)  "
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_SITUAC = 1  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Genera.MoveFirst
   
   Do While Not g_rst_Genera.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
   
      grd_Listad.Col = 0
      grd_Listad.Text = Format(g_rst_Genera!CajChc_CodCaj, "0000000000")
   
      If Trim(g_rst_Genera!CAJCHC_NUMDOC_2 & "") <> "" Then
         grd_Listad.Col = 1
         grd_Listad.Text = g_rst_Genera!CAJCHC_TIPDOC_2 & "-" & Trim(g_rst_Genera!CAJCHC_NUMDOC_2)
      End If
      
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(g_rst_Genera!MaePrv_RazSoc & "")
               
      grd_Listad.Col = 3
      grd_Listad.Text = gf_FormatoFecha(g_rst_Genera!CajChc_FecCaj)
   
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Genera!NOM_DEVREM & "")
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(moddat_g_str_DesMod) 'MONEDA
            
      grd_Listad.Col = 6
      grd_Listad.Text = Format(g_rst_Genera!CajChc_Import, "###,###,###,##0.00")
      
      grd_Listad.Col = 7
      grd_Listad.Text = Trim(g_rst_Genera!PROCESADO & "")
      
      grd_Listad.Col = 8
      grd_Listad.Text = "" 'Trim(g_rst_Genera!CAJDET_TipDoc & "")
      
      grd_Listad.Col = 9
      grd_Listad.Text = "" 'Trim(g_rst_Genera!CajDet_NumDoc & "")
         
      If Trim(g_rst_Genera!COMPAG_FECPAG & "") <> "" Then
         grd_Listad.Col = 10
         grd_Listad.Text = gf_FormatoFecha(g_rst_Genera!COMPAG_FECPAG)
      End If
         
      If Trim(g_rst_Genera!COMPAG_CODCOM & "") <> "" Then
         grd_Listad.Col = 11
         grd_Listad.Text = Format(g_rst_Genera!COMPAG_CODCOM, "00000000")
      End If
            
      grd_Listad.Col = 12
      grd_Listad.Text = g_rst_Genera!CAJCHC_TIPTAB '4=DEVOLUCION, 5=REEMBOLSO
            
      grd_Listad.Col = 13
      grd_Listad.Text = g_rst_Genera!CAJCHC_NUMERO 'ITEM (CORRELATIVO PK)
      
      grd_Listad.Col = 14
      grd_Listad.Text = Trim(g_rst_Genera!CAJCHC_DATCNT & "") 'ASIENTO CONTABLE
   
      '***AGREGAR AL ARREGLO
      ReDim Preserve l_arr_CajDet(UBound(l_arr_CajDet) + 1)
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_CodDet = g_rst_Genera!CajChc_CodCaj
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_CodCaj = g_rst_Genera!CajChc_CodCaj
      l_arr_CajDet(UBound(l_arr_CajDet)).CajChc_FecCaj = g_rst_Genera!CajChc_FecCaj
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_NumDoc = Trim(g_rst_Genera!CAJCHC_NUMDOC_2 & "")
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TipCpb_Lrg = g_rst_Genera!NOM_DEVREM
      
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TipDoc_Lrg = Trim(g_rst_Genera!TIPO_DOCUMENTO & "")
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_NumDoc = Trim(g_rst_Genera!CAJCHC_NUMDOC_2 & "")
      l_arr_CajDet(UBound(l_arr_CajDet)).MaePrv_RazSoc = Trim(g_rst_Genera!MaePrv_RazSoc & "")
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_Moneda = Trim(moddat_g_str_DesMod) 'MONEDA
      l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TotPpg = g_rst_Genera!CajChc_Import
      
      If g_rst_Genera!CAJCHC_TIPTAB = 5 Then
         'REEMBOLSO
         l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TipCpb = -3
         p_Importe = p_Importe - (g_rst_Genera!CajChc_Import)
      Else
         l_arr_CajDet(UBound(l_arr_CajDet)).CajDet_TipCpb = -2
         p_Importe = p_Importe + (g_rst_Genera!CajChc_Import)
      End If
   
      g_rst_Genera.MoveNext
   Loop
   
   Call gs_UbiIniGrid(grd_Listad)
   grd_Listad.Redraw = True
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
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
      .Cells(1, 2) = "Caja " & Format(moddat_g_str_Codigo, "0000000000")
      .Cells(1, 11) = "Fecha: " & moddat_g_str_FecIng & " "
      .Cells(2, 2) = "REPORTE DE ENTREGAS A RENDIR"
      .Range(.Cells(2, 2), .Cells(2, 11)).Merge
      .Range(.Cells(1, 2), .Cells(2, 11)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 11)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(3, 2), .Cells(3, 6)).Merge
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
      .Columns("B").ColumnWidth = 13 'codigo
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
          .Cells(r_int_NumFil + 2, 2) = "'" & Format(l_arr_CajDet(r_int_Contar).CajDet_CodDet, "0000000000") 'codigo
          If l_arr_CajDet(r_int_Contar).CajDet_FecEmi <> "" Then
             .Cells(r_int_NumFil + 2, 3) = "'" & gf_FormatoFecha(l_arr_CajDet(r_int_Contar).CajDet_FecEmi) 'fecha de emision
          End If
          .Cells(r_int_NumFil + 2, 4) = "'" & l_arr_CajDet(r_int_Contar).CajDet_TipCpb_Lrg 'tipo de comprobante
          .Cells(r_int_NumFil + 2, 5) = "'" & l_arr_CajDet(r_int_Contar).CajDet_Nserie     'serie
          .Cells(r_int_NumFil + 2, 6) = "'" & l_arr_CajDet(r_int_Contar).CajDet_NroCom     'numero
          .Cells(r_int_NumFil + 2, 7) = "'" & l_arr_CajDet(r_int_Contar).CajDet_NumDoc     'documento
          .Range(.Cells(r_int_NumFil + 2, 8), .Cells(r_int_NumFil + 2, 10)).Merge
          .Cells(r_int_NumFil + 2, 8) = "'" & l_arr_CajDet(r_int_Contar).MaePrv_RazSoc     'proveedor
          .Cells(r_int_NumFil + 2, 11) = l_arr_CajDet(r_int_Contar).CajDet_TotPpg
                                         
          If l_arr_CajDet(r_int_Contar).CajDet_TipCpb = 7 Or l_arr_CajDet(r_int_Contar).CajDet_TipCpb = 88 _
             Or l_arr_CajDet(r_int_Contar).CajDet_TipCpb = -3 Then
             r_dbl_MtoImp = r_dbl_MtoImp - l_arr_CajDet(r_int_Contar).CajDet_TotPpg
          Else
             r_dbl_MtoImp = r_dbl_MtoImp + l_arr_CajDet(r_int_Contar).CajDet_TotPpg
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

