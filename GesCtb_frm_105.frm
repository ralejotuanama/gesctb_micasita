VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Con_OpeFin_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10065
   ClientLeft      =   5025
   ClientTop       =   675
   ClientWidth     =   7920
   Icon            =   "GesCtb_frm_105.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10065
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7905
      _Version        =   65536
      _ExtentX        =   13944
      _ExtentY        =   17754
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   3105
         Left            =   30
         TabIndex        =   4
         Top             =   6900
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   5477
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
         Begin Threed.SSPanel pnl_Tit_NumMov 
            Height          =   285
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   5895
            _Version        =   65536
            _ExtentX        =   10398
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   5940
            TabIndex        =   6
            Top             =   60
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
            Height          =   1635
            Left            =   30
            TabIndex        =   2
            Top             =   360
            Width           =   7725
            _ExtentX        =   13626
            _ExtentY        =   2884
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Import 
            Height          =   315
            Left            =   5940
            TabIndex        =   7
            Top             =   2070
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ImpITF 
            Height          =   315
            Left            =   5940
            TabIndex        =   8
            Top             =   2400
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ImpTot 
            Height          =   315
            Left            =   5940
            TabIndex        =   9
            Top             =   2730
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
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
         Begin Threed.SSPanel pnl_PorITF 
            Height          =   315
            Left            =   4860
            TabIndex        =   10
            Top             =   2400
            Width           =   675
            _Version        =   65536
            _ExtentX        =   1191
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "(0.08%)"
            ForeColor       =   0
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
            Alignment       =   0
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   315
            Index           =   2
            Left            =   5430
            TabIndex        =   16
            Top             =   2730
            Width           =   465
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   315
            Index           =   1
            Left            =   5430
            TabIndex        =   15
            Top             =   2400
            Width           =   465
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   315
            Index           =   0
            Left            =   5430
            TabIndex        =   14
            Top             =   2070
            Width           =   465
         End
         Begin VB.Label Label10 
            Caption         =   "Importe:"
            Height          =   315
            Left            =   3810
            TabIndex        =   13
            Top             =   2070
            Width           =   885
         End
         Begin VB.Label Label11 
            Caption         =   "ITF:"
            Height          =   315
            Left            =   3810
            TabIndex        =   12
            Top             =   2400
            Width           =   795
         End
         Begin VB.Label Label12 
            Caption         =   "Importe Total:"
            Height          =   315
            Left            =   3810
            TabIndex        =   11
            Top             =   2730
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2085
         Left            =   30
         TabIndex        =   17
         Top             =   3300
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   3678
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
         Begin Threed.SSPanel pnl_CodBan 
            Height          =   315
            Left            =   1860
            TabIndex        =   18
            Top             =   60
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
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
         Begin Threed.SSPanel pnl_FecPag 
            Height          =   315
            Left            =   1860
            TabIndex        =   19
            Top             =   720
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
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
         End
         Begin Threed.SSPanel pnl_NumCom 
            Height          =   315
            Left            =   6360
            TabIndex        =   20
            Top             =   390
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_TipReg 
            Height          =   315
            Left            =   1860
            TabIndex        =   21
            Top             =   390
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_OfiPag 
            Height          =   315
            Left            =   1860
            TabIndex        =   22
            Top             =   1050
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
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
         Begin Threed.SSPanel pnl_ForPag 
            Height          =   315
            Left            =   1860
            TabIndex        =   23
            Top             =   1380
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
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
         Begin Threed.SSPanel pnl_CanPag 
            Height          =   315
            Left            =   1860
            TabIndex        =   24
            Top             =   1710
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
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
         Begin VB.Label Label22 
            Caption         =   "Canal de Pago:"
            Height          =   315
            Left            =   90
            TabIndex        =   31
            Top             =   1710
            Width           =   1425
         End
         Begin VB.Label Label21 
            Caption         =   "Forma de Pago:"
            Height          =   315
            Left            =   90
            TabIndex        =   30
            Top             =   1380
            Width           =   1425
         End
         Begin VB.Label Label20 
            Caption         =   "Oficina de Pago:"
            Height          =   315
            Left            =   90
            TabIndex        =   29
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Registro:"
            Height          =   255
            Left            =   90
            TabIndex        =   28
            Top             =   390
            Width           =   1635
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Comprobante:"
            Height          =   255
            Left            =   4740
            TabIndex        =   27
            Top             =   390
            Width           =   1635
         End
         Begin VB.Label Label16 
            Caption         =   "Fecha de Pago:"
            Height          =   315
            Left            =   90
            TabIndex        =   26
            Top             =   720
            Width           =   1425
         End
         Begin VB.Label Label13 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   90
            TabIndex        =   25
            Top             =   60
            Width           =   1305
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1425
         Left            =   30
         TabIndex        =   32
         Top             =   5430
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   2514
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1860
            TabIndex        =   33
            Top             =   60
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_TipDoc 
            Height          =   315
            Left            =   1860
            TabIndex        =   34
            Top             =   390
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1860
            TabIndex        =   35
            Top             =   720
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   1860
            TabIndex        =   36
            Top             =   1050
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label9 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   90
            TabIndex        =   40
            Top             =   1050
            Width           =   1005
         End
         Begin VB.Label Label8 
            Caption         =   "Nombre Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   39
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Doc. Ide. Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   38
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Nro. Operac. Refer.:"
            Height          =   315
            Left            =   90
            TabIndex        =   37
            Top             =   60
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1755
         Left            =   30
         TabIndex        =   41
         Top             =   1500
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   3096
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
         Begin Threed.SSPanel pnl_NomUsu 
            Height          =   315
            Left            =   1860
            TabIndex        =   42
            Top             =   1380
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
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
         Begin Threed.SSPanel pnl_HorMov 
            Height          =   315
            Left            =   6360
            TabIndex        =   43
            Top             =   720
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_NumMov 
            Height          =   315
            Left            =   1860
            TabIndex        =   44
            Top             =   390
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
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
         End
         Begin Threed.SSPanel pnl_TipMov 
            Height          =   315
            Left            =   1860
            TabIndex        =   45
            Top             =   1050
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
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
         Begin Threed.SSPanel pnl_FecMov 
            Height          =   315
            Left            =   1860
            TabIndex        =   46
            Top             =   720
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1879
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
         End
         Begin Threed.SSPanel pnl_SucMov 
            Height          =   315
            Left            =   1860
            TabIndex        =   47
            Top             =   60
            Width           =   5865
            _Version        =   65536
            _ExtentX        =   10345
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
            Caption         =   "Tipo Movimiento"
            Height          =   315
            Left            =   90
            TabIndex        =   53
            Top             =   1050
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Nro. Movimiento:"
            Height          =   315
            Left            =   60
            TabIndex        =   52
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Hora de Movimiento:"
            Height          =   315
            Left            =   4740
            TabIndex        =   51
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario:"
            Height          =   315
            Left            =   90
            TabIndex        =   50
            Top             =   1380
            Width           =   1785
         End
         Begin VB.Label Label14 
            Caption         =   "Fecha de Movimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   49
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label23 
            Caption         =   "Sucursal Movimiento:"
            Height          =   315
            Left            =   60
            TabIndex        =   48
            Top             =   60
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   30
         TabIndex        =   54
         Top             =   30
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   1244
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   4830
            Top             =   210
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   585
            Left            =   630
            TabIndex        =   55
            Top             =   30
            Width           =   3765
            _Version        =   65536
            _ExtentX        =   6641
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Consulta de Operaciones Financieras"
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
         Begin Threed.SSPanel pnl_FlgRev 
            Height          =   555
            Left            =   5880
            TabIndex        =   56
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "OPERACION REVERSADA"
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Outline         =   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_105.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel15 
         Height          =   675
         Left            =   30
         TabIndex        =   57
         Top             =   780
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7170
            Picture         =   "GesCtb_frm_105.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpCom 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_105.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Imprimir Comprobante"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Con_OpeFin_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ImpCom_Click()
   Dim r_str_DocIde     As String
   Dim r_str_NomCli     As String
   Dim r_str_FecPag     As String
   Dim r_str_SimMon     As String
   Dim r_str_moneda     As String
   Dim r_str_NumSol     As String
   Dim r_str_NumOpe     As String
   Dim r_int_TipNum     As Integer
   Dim r_str_NomBan     As String
   Dim r_str_NumCta     As String
   Dim r_str_CodPrd     As String
   Dim r_str_CodSub     As String
   Dim r_str_NomGas     As String
   Dim r_int_ConLin     As Integer
   Dim r_str_RevMov     As String
   Dim r_rst_Genera     As ADODB.Recordset

   If MsgBox("¿Está seguro de Imprimir el Comprobante?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Borrar Spool de PC (Cabecera)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_COMPGC WHERE "
   g_str_Parame = g_str_Parame & "COMPGC_CODTER = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   'Borrar Spool de PC (Detalle)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_COMPGD WHERE "
   g_str_Parame = g_str_Parame & "COMPGD_CODTER = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Call opecaj_gs_ComPago(moddat_g_str_CodGrp, opecaj_g_str_NumMov, opecaj_g_str_FecMov, 1, 1)
   
   Screen.MousePointer = 0
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat

   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "RPT_COMPGC"
   crp_Imprim.DataFiles(1) = "RPT_COMPGD"

   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_COMPAG_01.RPT"
   
   crp_Imprim.SelectionFormula = "{RPT_COMPGC.COMPGC_CODTER} = '" & modgen_g_str_NombPC & "'"
   
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_str_HorMov     As String
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicio
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Buscar()
   Dim r_int_TipMov     As Integer
   Dim r_str_CodPrd     As String
   Dim r_str_CodSub     As String
   Dim r_str_NumSol     As String
   Dim r_rst_Genera     As ADODB.Recordset
   
   Call gs_LimpiaGrid(grd_Listad)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & moddat_g_str_CodGrp & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMMOV = " & opecaj_g_str_NumMov & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV = " & opecaj_g_str_FecMov & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_SucMov.Caption = moddat_gf_ConsultaSucAge(moddat_g_str_Codigo, moddat_g_str_CodGrp)
      pnl_FecMov.Caption = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))
      pnl_HorMov.Caption = gf_FormatoHora(CStr(g_rst_Princi!CAJMOV_HORMOV))
      pnl_NumMov.Caption = Mid(CStr(g_rst_Princi!CAJMOV_FECMOV), 3, 2) & Format(g_rst_Princi!CAJMOV_NUMMOV, "00000")
      pnl_TipMov.Caption = CStr(g_rst_Princi!CAJMOV_TIPMOV) & " - " & moddat_gf_Consulta_ParDes("301", CStr(g_rst_Princi!CAJMOV_TIPMOV))
      pnl_NomUsu.Caption = Trim(g_rst_Princi!CAJMOV_USUMOV)
      
      r_int_TipMov = g_rst_Princi!CAJMOV_TIPMOV
         
      If Len(Trim(g_rst_Princi!CAJMOV_CODBAN)) > 0 And g_rst_Princi!CAJMOV_CODBAN <> "000000" Then
         pnl_CodBan.Caption = moddat_gf_Consulta_ParDes("505", Trim(g_rst_Princi!CAJMOV_CODBAN)) & " (Nro. Cuenta: " & Trim(g_rst_Princi!CAJMOV_NUMCTA & "") & ")"
      End If
      
      If g_rst_Princi!CAJMOV_FECDEP > 0 Then
         pnl_FecPag.Caption = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECDEP))
      End If
      
      'pnl_NumCta.Caption = Trim(g_rst_Princi!CAJMOV_NUMCTA & "")
      pnl_NumCom.Caption = Trim(g_rst_Princi!CAJMOV_NUMCOM & "")
         
      pnl_TipReg.Caption = moddat_gf_Consulta_ParDes("239", CStr(g_rst_Princi!CAJMOV_TIPREG))
         
      If g_rst_Princi!CAJMOV_TIPREG = 2 Then
         pnl_FecPag.Caption = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECREC))
         
         pnl_OfiPag.Caption = Trim(g_rst_Princi!CAJMOV_OFIPAG & "")
         pnl_ForPag.Caption = Trim(g_rst_Princi!CAJMOV_FORPAG & "")
         pnl_CanPag.Caption = Trim(g_rst_Princi!CAJMOV_CANPAG & "")
      End If
      
      If g_rst_Princi!CAJMOV_TIPMOV = 1101 Or g_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         pnl_NumOpe.Caption = gf_Formato_NumSol(Trim(g_rst_Princi!CAJMOV_NUMOPE & ""))
      Else
         pnl_NumOpe.Caption = gf_Formato_NumOpe(Trim(g_rst_Princi!CAJMOV_NUMOPE & "")) & IIf(Len(Trim(g_rst_Princi!CAJMOV_CODITE & "")) > 0, " (" & Trim(g_rst_Princi!CAJMOV_CODITE & "") & ")", "")
      End If
         
      pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!CAJMOV_TIPDOC)) & " - " & Trim(g_rst_Princi!CAJMOV_NUMDOC)
      
      If g_rst_Princi!CAJMOV_TIPDOC <> 7 Then
         pnl_NomCli.Caption = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!CAJMOV_TIPDOC), Trim(g_rst_Princi!CAJMOV_NUMDOC))
      End If
      
      pnl_Moneda.Caption = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!CAJMOV_MONPAG))
      pnl_Import.Caption = Format(g_rst_Princi!CAJMOV_IMPPAG, "###,###,##0.00") & " "
      pnl_ImpITF.Caption = Format(g_rst_Princi!CAJMOV_ITFIMP, "###,###,##0.00") & " "
      pnl_PorITF.Caption = "(" & Format(g_rst_Princi!CAJMOV_ITFPOR, "##0.00") & "%)"
      pnl_ImpTot.Caption = Format(g_rst_Princi!CAJMOV_IMPTOT, "###,###,##0.00") & " "
      
      lbl_SimMon(0).Caption = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!CAJMOV_MONPAG))
      lbl_SimMon(1).Caption = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!CAJMOV_MONPAG))
      lbl_SimMon(2).Caption = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!CAJMOV_MONPAG))
         
      If g_rst_Princi!CAJMOV_FLGREV = 0 Then
         pnl_FlgRev.Visible = False
      Else
         pnl_FlgRev.Visible = True
      End If
      
      Select Case r_int_TipMov
         Case "1101"    'Pago de Gastos de Cierre
            r_str_NumSol = Trim(g_rst_Princi!CAJMOV_NUMOPE & "")
         
            g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
            g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & r_str_NumSol & "' "
         
            If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
                Exit Sub
            End If
            
            r_rst_Genera.MoveFirst
            
            r_str_CodPrd = r_rst_Genera!SOLMAE_CODPRD
            r_str_CodSub = r_rst_Genera!SOLMAE_CODSUB
            
            r_rst_Genera.Close
            Set r_rst_Genera = Nothing
            
            
            
            'Buscar en Tabla de Gastos de Cierre
            g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
            g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & r_str_NumSol & "' "
         
            If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
                Exit Sub
            End If
            
            If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
               r_rst_Genera.MoveFirst
               
               Do While Not r_rst_Genera.EOF
                  grd_Listad.Rows = grd_Listad.Rows + 1
                  grd_Listad.Row = grd_Listad.Rows - 1
                  
                  'Buscando Descripción de Gastos Administrativos
                  grd_Listad.Col = 0
                  If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), r_str_CodPrd, r_str_CodSub, "007", Format(r_rst_Genera!GASADM_CODGAS, "00") & Format(r_rst_Genera!GASADM_TIPMON, "0")) Then
                     grd_Listad.Text = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
                  End If
                  
                  'Importe
                  grd_Listad.Col = 1
                  grd_Listad.Text = Format(r_rst_Genera!GASADM_IMPORT, "###,###,##0.00")
                  
                  r_rst_Genera.MoveNext
               Loop
            End If
            
            r_rst_Genera.Close
            Set r_rst_Genera = Nothing
            
         Case "1102"
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:        grd_Listad.Text = "CAPITAL"
            grd_Listad.Col = 1:        grd_Listad.Text = Format(g_rst_Princi!CAJMOV_CAPITA, "###,###,##0.00")
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = "INTERES"
            grd_Listad.Col = 1:     grd_Listad.Text = Format(g_rst_Princi!CAJMOV_INTERE, "###,###,##0.00")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = "SEGURO DESGRAVAMEN"
            grd_Listad.Col = 1:     grd_Listad.Text = Format(g_rst_Princi!CAJMOV_SEGDES, "###,###,##0.00")
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = "SEGURO INMUEBLE"
            grd_Listad.Col = 1:     grd_Listad.Text = Format(g_rst_Princi!CAJMOV_SEGVIV, "###,###,##0.00")
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = "PORTES"
            grd_Listad.Col = 1:     grd_Listad.Text = Format(g_rst_Princi!CAJMOV_OTRCAR, "###,###,##0.00")
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = "INTERES MORATORIO"
            grd_Listad.Col = 1:     grd_Listad.Text = Format(g_rst_Princi!CAJMOV_INTMOR, "###,###,##0.00")
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = "INTERES COMPENSATORIO"
            grd_Listad.Col = 1:     grd_Listad.Text = Format(g_rst_Princi!CAJMOV_INTCOM, "###,###,##0.00")
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = "GASTOS DE COBRANZA"
            grd_Listad.Col = 1:     grd_Listad.Text = Format(g_rst_Princi!CAJMOV_GASCOB, "###,###,##0.00")
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = "OTROS GASTOS"
            grd_Listad.Col = 1:     grd_Listad.Text = Format(g_rst_Princi!CAJMOV_OTRGAS, "###,###,##0.00")
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = "CAPITAL (PBP)"
            grd_Listad.Col = 1:     grd_Listad.Text = Format(g_rst_Princi!CAJMOV_CAPBBP, "###,###,##0.00")
         
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0:     grd_Listad.Text = "INTERES (PBP)"
            grd_Listad.Col = 1:     grd_Listad.Text = Format(g_rst_Princi!CAJMOV_INTBBP, "###,###,##0.00")
         
         Case "1103"
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0
            grd_Listad.Text = "DESEMBOLSO"
            
            grd_Listad.Col = 1
            grd_Listad.Text = Format(g_rst_Princi!CAJMOV_IMPPAG, "###,###,##0.00")
            
      End Select
      
      If grd_Listad.Rows > 0 Then
         Call gs_UbiIniGrid(grd_Listad)
      End If
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Inicio()
   grd_Listad.ColWidth(0) = 5895
   grd_Listad.ColWidth(1) = 1455
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignRightCenter
End Sub


