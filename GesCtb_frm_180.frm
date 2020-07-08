VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RegDes_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11595
   Icon            =   "GesCtb_frm_180.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel111 
      Height          =   9470
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11610
      _Version        =   65536
      _ExtentX        =   20479
      _ExtentY        =   16704
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   30
         TabIndex        =   13
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_180.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Rechazar 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_180.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Rechazar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10875
            Picture         =   "GesCtb_frm_180.frx":0890
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Aprobar 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_180.frx":0CD2
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Aprobar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Height          =   555
            Left            =   690
            TabIndex        =   17
            Top             =   30
            Width           =   7275
            _Version        =   65536
            _ExtentX        =   12832
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Aprobar Operaciones"
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
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   10920
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10350
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   80
            Picture         =   "GesCtb_frm_180.frx":0FDC
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   5925
         Left            =   30
         TabIndex        =   18
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   10451
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
         Begin TabDlg.SSTab SSTab1 
            Height          =   5805
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   11400
            _ExtentX        =   20108
            _ExtentY        =   10239
            _Version        =   393216
            Style           =   1
            Tabs            =   8
            Tab             =   6
            TabsPerRow      =   8
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "GesCtb_frm_180.frx":12E6
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Inmueble"
            TabPicture(1)   =   "GesCtb_frm_180.frx":1302
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Crédito"
            TabPicture(2)   =   "GesCtb_frm_180.frx":131E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(4)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Desembolso"
            TabPicture(3)   =   "GesCtb_frm_180.frx":133A
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).Control(1)=   "txt_ObsDes"
            Tab(3).ControlCount=   2
            TabCaption(4)   =   "Informe Legal"
            TabPicture(4)   =   "GesCtb_frm_180.frx":1356
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "txt_InfLeg"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Ev. Legal"
            TabPicture(5)   =   "GesCtb_frm_180.frx":1372
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "Label5"
            Tab(5).Control(1)=   "Label7"
            Tab(5).Control(2)=   "grd_Listad(6)"
            Tab(5).Control(3)=   "txt_ComCre"
            Tab(5).ControlCount=   4
            TabCaption(6)   =   "Datos del Desembolso"
            TabPicture(6)   =   "GesCtb_frm_180.frx":138E
            Tab(6).ControlEnabled=   -1  'True
            Tab(6).Control(0)=   "Label11"
            Tab(6).Control(0).Enabled=   0   'False
            Tab(6).Control(1)=   "SSPanel21"
            Tab(6).Control(1).Enabled=   0   'False
            Tab(6).Control(2)=   "SSPanel22"
            Tab(6).Control(2).Enabled=   0   'False
            Tab(6).Control(3)=   "SSPanel15"
            Tab(6).Control(3).Enabled=   0   'False
            Tab(6).Control(4)=   "pnl_Prycto_Dsm"
            Tab(6).Control(4).Enabled=   0   'False
            Tab(6).Control(5)=   "cmd_Dsm_ExpExc"
            Tab(6).Control(5).Enabled=   0   'False
            Tab(6).Control(6)=   "cmd_Dsm_ExpArc"
            Tab(6).Control(6).Enabled=   0   'False
            Tab(6).Control(7)=   "cmd_Cheque"
            Tab(6).Control(7).Enabled=   0   'False
            Tab(6).ControlCount=   8
            TabCaption(7)   =   "Seguimientos"
            TabPicture(7)   =   "GesCtb_frm_180.frx":13AA
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "grd_Listad(7)"
            Tab(7).ControlCount=   1
            Begin VB.CommandButton cmd_Cheque 
               Height          =   510
               Left            =   9510
               Picture         =   "GesCtb_frm_180.frx":13C6
               Style           =   1  'Graphical
               TabIndex        =   76
               ToolTipText     =   "Impresión de Cheque"
               Top             =   330
               Width           =   585
            End
            Begin VB.CommandButton cmd_Dsm_ExpArc 
               Height          =   510
               Left            =   10110
               Picture         =   "GesCtb_frm_180.frx":16D0
               Style           =   1  'Graphical
               TabIndex        =   70
               ToolTipText     =   "Exportar a Excel"
               Top             =   330
               Width           =   585
            End
            Begin VB.CommandButton cmd_Dsm_ExpExc 
               Height          =   510
               Left            =   10710
               Picture         =   "GesCtb_frm_180.frx":19DA
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Exportar a Excel"
               Top             =   330
               Width           =   585
            End
            Begin VB.TextBox txt_ComCre 
               Height          =   705
               Left            =   -74970
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   42
               Top             =   660
               Width           =   11235
            End
            Begin VB.TextBox txt_InfLeg 
               Height          =   5475
               Left            =   -74910
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   41
               Top             =   390
               Width           =   11200
            End
            Begin VB.TextBox txt_ObsDes 
               Height          =   975
               Left            =   -74910
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   4920
               Width           =   11235
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   5535
               Index           =   0
               Left            =   -74910
               TabIndex        =   23
               Top             =   390
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   9763
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   5535
               Index           =   1
               Left            =   -74910
               TabIndex        =   24
               Top             =   390
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   9763
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   5535
               Index           =   4
               Left            =   -74910
               TabIndex        =   25
               Top             =   390
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   9763
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4485
               Index           =   3
               Left            =   -74910
               TabIndex        =   26
               Top             =   405
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   7911
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4275
               Index           =   6
               Left            =   -75000
               TabIndex        =   43
               Top             =   1650
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   7541
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   5535
               Index           =   7
               Left            =   -74910
               TabIndex        =   46
               Top             =   390
               Width           =   11250
               _ExtentX        =   19844
               _ExtentY        =   9763
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_Prycto_Dsm 
               Height          =   315
               Left            =   1680
               TabIndex        =   47
               Top             =   480
               Width           =   6660
               _Version        =   65536
               _ExtentX        =   11747
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   645
               Left            =   60
               TabIndex        =   49
               Top             =   2895
               Width           =   11235
               _Version        =   65536
               _ExtentX        =   19817
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
               Begin VB.CommandButton cmd_Dsm_Editar 
                  Height          =   585
                  Left            =   10590
                  Picture         =   "GesCtb_frm_180.frx":1CE4
                  Style           =   1  'Graphical
                  TabIndex        =   10
                  ToolTipText     =   "Modificar Registro"
                  Top             =   40
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Dsm_Nuevo 
                  Height          =   585
                  Left            =   9390
                  Picture         =   "GesCtb_frm_180.frx":1FEE
                  Style           =   1  'Graphical
                  TabIndex        =   0
                  ToolTipText     =   "Adicionar Registro"
                  Top             =   40
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Dsm_Borrar 
                  Height          =   585
                  Left            =   9990
                  Picture         =   "GesCtb_frm_180.frx":22F8
                  Style           =   1  'Graphical
                  TabIndex        =   9
                  ToolTipText     =   "Eliminar Registro"
                  Top             =   40
                  Width           =   585
               End
            End
            Begin Threed.SSPanel SSPanel22 
               Height          =   2010
               Left            =   60
               TabIndex        =   51
               Top             =   870
               Width           =   11250
               _Version        =   65536
               _ExtentX        =   19844
               _ExtentY        =   3545
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_Dsm 
                  Height          =   1650
                  Left            =   0
                  TabIndex        =   52
                  Top             =   15
                  Width           =   11235
                  _ExtentX        =   19817
                  _ExtentY        =   2910
                  _Version        =   393216
                  Rows            =   1
                  Cols            =   14
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  SelectionMode   =   1
                  Appearance      =   0
               End
               Begin Threed.SSPanel pnl_SumTot_Dsm 
                  Height          =   285
                  Left            =   6060
                  TabIndex        =   53
                  Top             =   1680
                  Width           =   1200
                  _Version        =   65536
                  _ExtentX        =   2117
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "9,999,999.99 "
                  ForeColor       =   16777215
                  BackColor       =   192
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
               Begin Threed.SSPanel pnl_TotPtmo_Dsm 
                  Height          =   285
                  Left            =   1095
                  TabIndex        =   54
                  Top             =   1680
                  Width           =   1200
                  _Version        =   65536
                  _ExtentX        =   2117
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "9,999,999.99 "
                  ForeColor       =   16777215
                  BackColor       =   192
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
               Begin VB.Label lbl_Bono_Dsm 
                  AutoSize        =   -1  'True
                  Caption         =   ".."
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   57
                  Top             =   1740
                  Width           =   90
               End
               Begin VB.Label lbl_Total 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Total ==> "
                  Height          =   195
                  Left            =   5325
                  TabIndex        =   56
                  Top             =   1740
                  Width           =   720
               End
               Begin VB.Label lbl_Totale 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Distribuir ==> "
                  Height          =   195
                  Index           =   1
                  Left            =   135
                  TabIndex        =   55
                  Top             =   1740
                  Width           =   960
               End
            End
            Begin Threed.SSPanel SSPanel21 
               Height          =   2205
               Left            =   45
               TabIndex        =   58
               Top             =   3540
               Width           =   11265
               _Version        =   65536
               _ExtentX        =   19870
               _ExtentY        =   3889
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
               BorderWidth     =   1
               BevelOuter      =   0
               BevelInner      =   1
               Begin VB.ComboBox cmb_TipMto_Dsm 
                  Height          =   315
                  Left            =   5880
                  Style           =   2  'Dropdown List
                  TabIndex        =   2
                  Top             =   120
                  Width           =   2985
               End
               Begin VB.TextBox txt_Descrp_Dsm 
                  Height          =   315
                  Left            =   1680
                  MaxLength       =   250
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   7
                  Top             =   1785
                  Width           =   7170
               End
               Begin VB.ComboBox cmb_NroCta_Dsm 
                  Height          =   315
                  Left            =   1680
                  Style           =   2  'Dropdown List
                  TabIndex        =   5
                  Top             =   1125
                  Width           =   3000
               End
               Begin VB.CommandButton cmd_Dsm_Insert 
                  Height          =   585
                  Left            =   10020
                  Picture         =   "GesCtb_frm_180.frx":2602
                  Style           =   1  'Graphical
                  TabIndex        =   8
                  Tag             =   "2"
                  Top             =   1560
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Dsm_Cancel 
                  Height          =   585
                  Left            =   10620
                  Picture         =   "GesCtb_frm_180.frx":290C
                  Style           =   1  'Graphical
                  TabIndex        =   11
                  Top             =   1560
                  Width           =   585
               End
               Begin VB.TextBox txt_ANombre_Dsm 
                  Height          =   315
                  Left            =   1680
                  MaxLength       =   250
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   6
                  Top             =   1455
                  Width           =   7170
               End
               Begin VB.ComboBox cmb_FrmDsm_Dsm 
                  Height          =   315
                  Left            =   1680
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   120
                  Width           =   3000
               End
               Begin VB.ComboBox cmb_EntFin_Dsm 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   1680
                  Style           =   2  'Dropdown List
                  TabIndex        =   4
                  Top             =   795
                  Width           =   7170
               End
               Begin EditLib.fpDoubleSingle ipp_Import_Dsm 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   3
                  Top             =   465
                  Width           =   3000
                  _Version        =   196608
                  _ExtentX        =   5292
                  _ExtentY        =   556
                  Enabled         =   -1  'True
                  MousePointer    =   0
                  Object.TabStop         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   1
                  ThreeDInsideHighlightColor=   -2147483637
                  ThreeDInsideShadowColor=   -2147483642
                  ThreeDInsideWidth=   1
                  ThreeDOutsideStyle=   1
                  ThreeDOutsideHighlightColor=   -2147483628
                  ThreeDOutsideShadowColor=   -2147483632
                  ThreeDOutsideWidth=   1
                  ThreeDFrameWidth=   0
                  BorderStyle     =   0
                  BorderColor     =   -2147483642
                  BorderWidth     =   1
                  ButtonDisable   =   0   'False
                  ButtonHide      =   0   'False
                  ButtonIncrement =   1
                  ButtonMin       =   0
                  ButtonMax       =   100
                  ButtonStyle     =   0
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   2
                  AlignTextV      =   0
                  AllowNull       =   0   'False
                  NoSpecialKeys   =   0
                  AutoAdvance     =   0   'False
                  AutoBeep        =   0   'False
                  CaretInsert     =   0
                  CaretOverWrite  =   3
                  UserEntry       =   0
                  HideSelection   =   -1  'True
                  InvalidColor    =   -2147483637
                  InvalidOption   =   0
                  MarginLeft      =   3
                  MarginTop       =   3
                  MarginRight     =   3
                  MarginBottom    =   3
                  NullColor       =   -2147483637
                  OnFocusAlignH   =   0
                  OnFocusAlignV   =   0
                  OnFocusNoSelect =   0   'False
                  OnFocusPosition =   0
                  ControlType     =   0
                  Text            =   "0.00"
                  DecimalPlaces   =   2
                  DecimalPoint    =   "."
                  FixedPoint      =   -1  'True
                  LeadZero        =   0
                  MaxValue        =   "9000000000"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
                  IncInt          =   1
                  IncDec          =   1
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   0
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin Threed.SSPanel pnl_Moneda_Dsm 
                  Height          =   315
                  Left            =   5880
                  TabIndex        =   59
                  Top             =   465
                  Width           =   2970
                  _Version        =   65536
                  _ExtentX        =   5239
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
               Begin Threed.SSPanel pnl_NroCCI_Dsm 
                  Height          =   315
                  Left            =   5880
                  TabIndex        =   60
                  Top             =   1125
                  Width           =   2970
                  _Version        =   65536
                  _ExtentX        =   5239
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
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo Monto:"
                  Height          =   195
                  Left            =   4980
                  TabIndex        =   69
                  Top             =   195
                  Width           =   855
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  Caption         =   "Descripción:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   68
                  Top             =   1845
                  Width           =   885
               End
               Begin VB.Label Label25 
                  AutoSize        =   -1  'True
                  Caption         =   "Importe Desembolso:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   67
                  Top             =   510
                  Width           =   1485
               End
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
                  Caption         =   "Forma Desembolso:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   66
                  Top             =   195
                  Width           =   1395
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "Nro Cuenta:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   65
                  Top             =   1185
                  Width           =   855
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Moneda:"
                  Height          =   195
                  Left            =   4980
                  TabIndex        =   64
                  Top             =   510
                  Width           =   630
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "A Nombre de:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   63
                  Top             =   1515
                  Width           =   975
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Entidad Financiera:"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   62
                  Top             =   855
                  Width           =   1365
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Nro CCI:"
                  Height          =   195
                  Left            =   4980
                  TabIndex        =   61
                  Top             =   1200
                  Width           =   600
               End
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Proyecto:"
               Height          =   195
               Left            =   240
               TabIndex        =   48
               Top             =   525
               Width           =   1275
            End
            Begin VB.Label Label7 
               Caption         =   "Comité de Créditos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -74970
               TabIndex        =   45
               Top             =   420
               Width           =   3495
            End
            Begin VB.Label Label5 
               Caption         =   "Datos de la Evaluación"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -74970
               TabIndex        =   44
               Top             =   1440
               Width           =   3495
            End
            Begin VB.Label Label6 
               Caption         =   "Observaciones"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -74970
               TabIndex        =   29
               Top             =   2160
               Width           =   2805
            End
            Begin VB.Label Label59 
               Caption         =   "Comité de Créditos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -74970
               TabIndex        =   28
               Top             =   360
               Width           =   2805
            End
            Begin VB.Label Label3 
               Caption         =   "Contratos y Bloqueo Registral"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   -74970
               TabIndex        =   27
               Top             =   1530
               Width           =   2805
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1185
         Left            =   30
         TabIndex        =   30
         Top             =   8220
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2090
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
         Begin VB.TextBox txt_NroDsm_Dsm 
            Height          =   315
            Left            =   1770
            MaxLength       =   250
            ScrollBars      =   2  'Vertical
            TabIndex        =   71
            Top             =   150
            Width           =   2895
         End
         Begin VB.TextBox txt_Comentario 
            Height          =   555
            Left            =   1770
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            Top             =   480
            Width           =   6930
         End
         Begin EditLib.fpDateTime ipp_FecDsm_Dsm 
            Height          =   315
            Left            =   6000
            TabIndex        =   72
            Top             =   150
            Width           =   1980
            _Version        =   196608
            _ExtentX        =   3492
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "24/04/2015"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label lbl_FchDsm_Dsm 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Reg:"
            Height          =   195
            Left            =   5130
            TabIndex        =   75
            Top             =   210
            Width           =   840
         End
         Begin VB.Label lbl_NumDsm_Dsm 
            AutoSize        =   -1  'True
            Caption         =   "Nro Transferencia:"
            Height          =   195
            Left            =   270
            TabIndex        =   73
            Top             =   210
            Width           =   1320
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Comentario:"
            Height          =   195
            Left            =   270
            TabIndex        =   31
            Top             =   660
            Width           =   840
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   32
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   930
            TabIndex        =   33
            Top             =   390
            Width           =   6435
            _Version        =   65536
            _ExtentX        =   11351
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   8730
            TabIndex        =   34
            Top             =   60
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
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
         End
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   930
            TabIndex        =   35
            Top             =   60
            Width           =   6435
            _Version        =   65536
            _ExtentX        =   11351
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
         Begin Threed.SSPanel pnl_EstadoActual 
            Height          =   315
            Left            =   8730
            TabIndex        =   36
            Top             =   390
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
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
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   150
            TabIndex        =   40
            Top             =   120
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro.Operación:"
            Height          =   195
            Left            =   7560
            TabIndex        =   39
            Top             =   120
            Width           =   1080
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Producto:"
            Height          =   195
            Left            =   150
            TabIndex        =   38
            Top             =   450
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Instancia:"
            Height          =   195
            Left            =   7560
            TabIndex        =   37
            Top             =   450
            Width           =   690
         End
      End
   End
End
Attribute VB_Name = "frm_RegDes_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_MtoHip     As Double
Dim l_str_MonBlq     As String
Dim l_dbl_ImpTas     As Double
Dim l_dbl_ImpNot     As Double
Dim l_dbl_ImpEst     As Double
Dim l_dbl_ImpEva     As Double
Dim l_dbl_ImpAdm     As Double
Dim l_dbl_ImpRed     As Double
Dim l_dbl_ImpBlq     As Double
Dim l_int_ChqReg     As Integer
Dim l_int_PolReg     As Integer
Dim l_int_FiaReg     As Integer
Dim l_int_CerReg     As Integer
Dim l_int_FlgCVt     As Integer
Dim l_int_MonCvt     As Integer
Dim l_str_CodMod     As String
Dim l_str_Moneda     As String
Dim l_dbl_ImpPtm     As Double
Dim l_str_Prmtor     As String
Dim l_str_CodBan     As String
Dim l_str_PryBan     As String
Dim l_str_DocPrm     As String
   
Dim l_arr_CtaBco()   As moddat_tpo_Genera
Dim l_arr_Bancos()   As moddat_tpo_Genera

Private Sub cmb_NroCta_Dsm_Click()
Dim r_int_Fila As Integer

    pnl_NroCCI_Dsm.Caption = ""
    For r_int_Fila = 1 To UBound(l_arr_CtaBco)
        If (l_arr_CtaBco(r_int_Fila).Genera_Codigo = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo) And _
            Trim(l_arr_CtaBco(r_int_Fila).Genera_Nombre) = Trim(cmb_NroCta_Dsm.Text)) Then
            pnl_NroCCI_Dsm.Caption = Trim(l_arr_CtaBco(r_int_Fila).Genera_Refere)
            
            If cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 2 Then 'transferencia
               txt_ANombre_Dsm.Text = ""
               txt_ANombre_Dsm.Text = Trim(l_arr_CtaBco(r_int_Fila).Genera_NomCli & "")
            End If
            Exit For
        End If
    Next
End Sub
 
Private Sub cmb_TipMto_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Import_Dsm)
   End If
End Sub

Private Sub cmd_Aprobar_Click()
Dim r_bol_Estado  As Boolean
Dim r_int_Fila    As Integer
Dim r_str_NumAsi  As String

   If Len(Trim(moddat_g_str_NumOpe)) = 0 Then
      MsgBox "Tiene que Haber un Nro. Operación.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Validando
   If (CDbl(pnl_SumTot_Dsm.Caption) <> CDbl(CStr(l_dbl_ImpPtm))) Then
      SSTab1.Tab = 6
      MsgBox "El préstamo total no es igual al distribuido en la pestaña datos del desembolso", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Validar Cuentas del Promotor
   'For r_int_fila = 0 To grd_Listad_Dsm.Rows - 1
   '    If (grd_Listad_Dsm.RowHeight(r_int_fila) > 0) Then
   '        If (Len(Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 8))) = 0) Or (Len(Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 9))) = 0) Then
   '            SSTab1.Tab = 6
   '            MsgBox "Faltan ingresar datos en algunos registros, de la pestaña datos del desembolso.", vbExclamation, modgen_g_str_NomPlt
   '            Exit Sub
   '        End If
   '    End If
   'Next
   
   If Len(Trim(txt_NroDsm_Dsm.Text)) = 0 Then
      SSTab1.Tab = 6
      MsgBox "Debe de ingresar el Nro Cheque o Transferencia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NroDsm_Dsm)
      Exit Sub
   End If
   
   If Len(Trim(ipp_FecDsm_Dsm.Text)) = 0 Then
      SSTab1.Tab = 6
      MsgBox "Debe de ingresar la fecha de registro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecDsm_Dsm)
      Exit Sub
   End If
   
'   If (Format(ipp_FecDsm_Dsm.Text, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
'       Format(ipp_FecDsm_Dsm.Text, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
'       MsgBox "Intenta Registrar un documento en un periodo cerrado.", vbExclamation, modgen_g_str_NomPlt
'       Call gs_SetFocus(ipp_FecDsm_Dsm)
'        Exit Sub
'   End If
   
   Screen.MousePointer = 11
   'descab_fecreg , descab_horreg
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DESCAB_NUMOPE, DESCAB_CODEST "
   g_str_Parame = g_str_Parame & "   FROM CRE_DESPROCAB "
   g_str_Parame = g_str_Parame & "  WHERE DESCAB_NUMOPE = '" & moddat_g_str_NumOpe & "'   "
   g_str_Parame = g_str_Parame & "    AND descab_fecreg = '" & moddat_g_str_FecRec & "'    "
   g_str_Parame = g_str_Parame & "    AND descab_horreg = '" & moddat_g_str_FecHip & "'    "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   If (Trim(CStr(g_rst_Princi!DESCAB_CODEST)) <> Trim(moddat_g_str_CodIte)) Then
      MsgBox "Este registro ya ha cambiado de estado.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
            
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Actualizando
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
   
      g_str_Parame = "usp_Actualiza_cre_desprocab("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecRec & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecHip & "', "
      g_str_Parame = g_str_Parame & "'4', " 'aprobar
      g_str_Parame = g_str_Parame & "'6', "
   
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_FESLNT
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLEG
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_FEENNT
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLE2
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_FERELG
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNOPE
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_FERECE
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNOP2
      g_str_Parame = g_str_Parame & "'" & txt_Comentario.Text & "', "
      '--------------------------------------------------------------
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      Screen.MousePointer = 0
   Loop
   
   r_bol_Estado = fs_guardar_ctaPromotor
   If (r_bol_Estado = False) Then
       Exit Sub
   End If
   
   'Aprobar Operaciones 2DA Parte
    g_str_Parame = "usp_Actualiza_cre_desprocab("
    g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
    g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecRec & "', "
    g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecHip & "', "
    g_str_Parame = g_str_Parame & "'5', "
    g_str_Parame = g_str_Parame & "'8', "
    g_str_Parame = g_str_Parame & "'', " 'fecha solicitud notaria
    g_str_Parame = g_str_Parame & "'', " 'comentario legal 1
    g_str_Parame = g_str_Parame & "'', " 'fecha entrega notaria
    g_str_Parame = g_str_Parame & "'', " 'comentario legal 2
    g_str_Parame = g_str_Parame & "'', " 'fecha de recepcion 2
    g_str_Parame = g_str_Parame & "'', "
    g_str_Parame = g_str_Parame & ", " '& Format(ipp_FecConstancia.Text, "yyyymmdd") & "', "
    g_str_Parame = g_str_Parame & "'', " ' & txt_Comentario.Text & "', "
    g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLE2
    '------------
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
    g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
    g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', 1) "
      
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
       Exit Sub
    End If
    
   'Final Operaciones 2DA Parte
   r_str_NumAsi = ""
   r_int_Fila = 0
   r_bol_Estado = True
   
   For r_int_Fila = 1 To grd_Listad_Dsm.Rows - 1
       If grd_Listad_Dsm.TextMatrix(r_int_Fila, 0) = 1 Then 'FORMA DESEMBOLSO(Cheque Simple)
          r_bol_Estado = False
          Exit For
       End If
   Next
   
   If r_bol_Estado = True Then
      Call fs_GeneraAsiento(r_str_NumAsi)
   End If
   
   'Enviando Correo Electrónico
   Call fs_Envia_Correo("APROBACION")
   
   'Imprime liquidacion
   Screen.MousePointer = 0
   If (moddat_g_int_CntErr = 0) Then
       If r_bol_Estado = True Then
          MsgBox "El proceso se grabó exitosamente." & vbCrLf & "El asiento generado es:" & r_str_NumAsi, vbInformation, modgen_g_str_NomPlt
       Else
          MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
       End If
       frm_RegDes_01.fs_Buscar_Creditos
       Unload Me
   End If
End Sub

Private Sub fs_Envia_Correo(p_Estado As String)

   modgen_g_str_Mail_Asunto = "PAGO PROMOTOR - AREA TESORERIA - " & p_Estado & " (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & moddat_g_str_NumSol & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE OPERACION : " & moddat_g_str_NumOpe & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   
   Call fs_Envia_Correo_Prom(mps_Sesion, mps_Mensaj, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, "", "", False, True, False, False, True, True)
End Sub

Private Sub fs_GeneraAsiento(ByRef p_NumAsi As String)
Dim r_arr_LogPro()  As modprc_g_tpo_LogPro
Dim r_int_NumIte    As Integer
Dim r_int_NumAsi    As Integer
Dim r_str_Glosa     As String
Dim r_dbl_MtoSol    As Double
Dim r_dbl_MtoDol    As Double
Dim r_str_FechaL    As String
Dim r_str_FechaC    As String
Dim r_int_NumLib    As Integer
Dim r_str_Origen    As String
Dim r_int_Contar    As Integer
Dim r_str_CtaHab    As String
Dim r_str_CtaDeb    As String
Dim r_dbl_TipSbs    As Double
Dim r_str_TipNot    As String
Dim r_dbl_SumImp    As Double
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "B"
   r_int_NumLib = 12
   
   r_int_NumAsi = 0
   r_int_NumIte = 0
   r_str_FechaC = Format(ipp_FecDsm_Dsm.Text, "yyyymmdd")
   r_str_FechaL = ipp_FecDsm_Dsm.Text
   
   r_int_PerAno = modctb_int_PerAno 'Year(ipp_FecDsm_Dsm.Text)
   r_int_PerMes = modctb_int_PerMes 'Month(ipp_FecDsm_Dsm.Text)
      
   p_NumAsi = ""
   
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
   p_NumAsi = CStr(r_int_NumAsi)
   
   r_str_Glosa = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & "/" & fs_ExtraeApellido(moddat_g_str_NomCli) & "/" & l_str_DocPrm
   r_str_Glosa = Mid(r_str_Glosa, 1, 60)
   
   'Insertar en CABECERA
   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                 r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FechaL, "1")
   r_dbl_SumImp = 0
   r_int_NumIte = 1
   For r_int_Contar = 1 To grd_Listad_Dsm.Rows - 1
       Select Case CInt(grd_Listad_Dsm.TextMatrix(r_int_Contar, 2))
              Case 1: r_str_CtaDeb = "251419010101" 'DESEMBOLSO
              Case 2: r_str_CtaDeb = "291807010114" 'AFP
              Case 3: r_str_CtaDeb = "291807010113" 'BONO
              Case 4: r_str_CtaDeb = "291807010115" 'BMS
              Case 5: r_str_CtaDeb = "291807010113" 'BONO MEF
       End Select
   
       r_dbl_MtoSol = CDbl(grd_Listad_Dsm.TextMatrix(r_int_Contar, 4))
       r_dbl_MtoDol = 0
       r_dbl_SumImp = r_dbl_SumImp + CDbl(grd_Listad_Dsm.TextMatrix(r_int_Contar, 4))
       
       Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                            r_int_NumAsi, r_int_NumIte, r_str_CtaDeb, CDate(r_str_FechaL), _
                                            r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
       r_int_NumIte = r_int_NumIte + 1
   Next
   
   r_str_CtaHab = "111301060102"
   r_dbl_MtoSol = r_dbl_SumImp
   r_dbl_MtoDol = 0
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, r_int_NumIte, r_str_CtaHab, CDate(r_str_FechaL), _
                                        r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
   
   'MsgBox "Asiento contable generado : " & r_int_NumAsi & vbCrLf & _
   '       "Nro de Items Generados: " & r_int_NumIte - 1, vbInformation, modgen_g_str_NomPlt
End Sub

Function fs_ExtraeApellido(p_Apellido As String) As String
Dim r_int_PosIni As Integer
Dim r_int_PosFin As Integer

   r_int_PosIni = 0
   r_int_PosIni = InStr(1, Trim(p_Apellido), " ") + 1
   r_int_PosFin = InStr(r_int_PosIni, Trim(p_Apellido), " ")
   
   If r_int_PosFin = 0 And r_int_PosIni - 1 > 0 Then
      fs_ExtraeApellido = Trim(Mid(Trim(p_Apellido), 1, r_int_PosIni - 1))
   ElseIf r_int_PosFin > 0 Then
      fs_ExtraeApellido = Trim(Mid(Trim(p_Apellido), 1, r_int_PosFin))
   Else
      fs_ExtraeApellido = Trim(p_Apellido)
   End If
End Function

Private Sub cmd_Cheque_Click()
Dim r_str_CadAux   As String

   If grd_Listad_Dsm.Rows <= 1 Then
      Exit Sub
   End If

   If grd_Listad_Dsm.TextMatrix(1, 0) = 1 Then
      'CHEQUE
      r_str_CadAux = ""
      
      frm_Ctb_PagCom_08.ipp_FecChq.Text = date
      frm_Ctb_PagCom_08.txt_NomDe.Text = Trim(grd_Listad_Dsm.TextMatrix(1, 8)) 'A NOMBRE DE
      
      frm_Ctb_PagCom_08.pnl_Import.Caption = Trim(pnl_SumTot_Dsm.Caption) & " "  'IMPORTE
      
      frm_Ctb_PagCom_08.pnl_Moneda.Caption = l_str_Moneda 'MONEDA
      frm_Ctb_PagCom_08.txt_CodOrigen.Text = "MODULO_DESEMBOLSO_PROMOTOR"
      frm_Ctb_PagCom_08.txt_CodOrigen.Tag = pnl_NumOpe.Caption 'CODIGO
      frm_Ctb_PagCom_08.fs_NumeroLetra
      frm_Ctb_PagCom_08.Show 1
   Else
      MsgBox "Solo se emiten, forma de desembolso tipo cheques simple.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_Dsm_ExpArc_Click()
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer
Dim r_int_NumRes    As Integer
Dim r_str_NomRes    As String
Dim r_str_Cadena    As String
Dim r_str_CadAux    As String
Dim r_dbl_PlaTot    As Double
Dim r_int_RegTot    As Integer
Dim r_int_Contad    As Integer
Dim r_int_PosIni    As Integer
Dim r_str_AuxRuc    As String

   r_int_PerAno = Year(moddat_g_str_FecSis)
   r_int_PerMes = Month(moddat_g_str_FecSis)
   r_str_NomRes = moddat_g_str_RutLoc & "\" & Format(moddat_g_str_FecSis, "yyyymm") & "_Transferencias.TXT"
                      
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
      
   r_str_Cadena = ""
   r_dbl_PlaTot = 0
       
   r_str_CadAux = ""
   For r_int_Contad = 1 To 68
       r_str_CadAux = r_str_CadAux & " "
   Next
   
   r_int_RegTot = 0
   For r_int_Contad = 1 To grd_Listad_Dsm.Rows - 1
       If grd_Listad_Dsm.TextMatrix(r_int_Contad, 0) = 2 Then
          r_dbl_PlaTot = r_dbl_PlaTot + CDbl(grd_Listad_Dsm.TextMatrix(r_int_Contad, 4))
          r_int_RegTot = r_int_RegTot + 1
       End If
   Next
   
   r_str_Cadena = r_str_Cadena & "50000110661000100040896PEN" & Format(r_dbl_PlaTot * 100, "000000000000000") & _
                                 "A" & Format(moddat_g_str_FecSis, "yyyymmdd") & "H" & "TRANSFERENCIAS           " & _
                                 Format(r_int_RegTot, "000000") & "N" & r_str_CadAux
   Print #r_int_NumRes, r_str_Cadena
   
   r_str_CadAux = ""
   For r_int_Contad = 1 To 101
       r_str_CadAux = r_str_CadAux & " "
   Next
   Dim r_str_TipDoc As String
   Dim r_str_TipCta As String
   Dim r_str_NumCta As String
   
   r_int_PosIni = InStr(1, Trim(l_str_DocPrm), "-")
   r_str_AuxRuc = CLng(Mid(Trim(l_str_DocPrm), 1, r_int_PosIni - 1))
   Select Case CLng(Trim(r_str_AuxRuc))
          Case 1
               r_str_TipDoc = "L"
          Case 4
               r_str_TipDoc = "E"
          Case 7
              r_str_TipDoc = "R"
   End Select
   
   For r_int_Contad = 1 To grd_Listad_Dsm.Rows - 1
       If grd_Listad_Dsm.TextMatrix(r_int_Contad, 0) = 2 Then
          r_int_PosIni = 0
          r_str_AuxRuc = ""
          If CLng(grd_Listad_Dsm.TextMatrix(r_int_Contad, 5)) = 2 Then
             r_str_TipCta = "P"
             r_str_NumCta = Mid(Trim(grd_Listad_Dsm.TextMatrix(r_int_Contad, 7)), 1, 8) & "00" & Mid(Trim(grd_Listad_Dsm.TextMatrix(r_int_Contad, 7)), 9, 10)
          Else
             r_str_TipCta = "I"
             r_str_NumCta = fs_numCCI(grd_Listad_Dsm.TextMatrix(r_int_Contad, 5), grd_Listad_Dsm.TextMatrix(r_int_Contad, 7))
          End If
          r_str_Cadena = ""
          
          r_int_PosIni = InStr(1, Trim(l_str_DocPrm), "-")
          r_str_AuxRuc = Trim(Mid(Trim(l_str_DocPrm), r_int_PosIni + 1, 30))
          
          r_str_Cadena = r_str_Cadena & "002" & r_str_TipDoc & _
                                        Left(Trim(r_str_AuxRuc) & "            ", 12) & _
                                        r_str_TipCta & r_str_NumCta & _
                                        Left(Trim(grd_Listad_Dsm.TextMatrix(r_int_Contad, 8)) & "                                        ", 40) & _
                                        Format(CDbl(grd_Listad_Dsm.TextMatrix(r_int_Contad, 4)) * 100, "000000000000000") & _
                                        Left(Trim(gf_Formato_NumSol(moddat_g_str_NumSol)) & "                                        ", 40) & _
                                        r_str_CadAux
          Print #r_int_NumRes, r_str_Cadena
       End If
   Next
               
   'Cerrando Archivo Resumen
   Close r_int_NumRes
   
   '-----------MENSAJE FINAL------------------------------------------
   MsgBox "Archivo generado con éxito: " & r_str_NomRes, vbInformation, modgen_g_str_NomPlt
End Sub

Private Function fs_numCCI(ByVal p_numBco As String, ByVal p_NumCta As String) As String
Dim r_int_Fila   As Integer
    fs_numCCI = "                    "
    For r_int_Fila = 1 To UBound(l_arr_CtaBco)
        If (CLng(l_arr_CtaBco(r_int_Fila).Genera_Codigo) = CLng(p_numBco) And _
            Trim(l_arr_CtaBco(r_int_Fila).Genera_Nombre) = Trim(p_NumCta)) Then
            fs_numCCI = Left(Trim(l_arr_CtaBco(r_int_Fila).Genera_Refere) & "                    ", 20)
            Exit For
        End If
    Next
End Function

Private Sub cmd_Grabar_Click()
Dim r_bol_Estado  As Boolean

    If Len(Trim(moddat_g_str_NumOpe)) = 0 Then
       MsgBox "Tiene que Haber un Nro. Operación.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
    End If
            
    'Validando
    If (CDbl(pnl_SumTot_Dsm.Caption) > CDbl(CStr(l_dbl_ImpPtm))) Then
       SSTab1.Tab = 6
       MsgBox "El préstamo total no es igual al distribuido en la pestaña datos del desembolso", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
    End If
    
    If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    moddat_g_int_FlgGOK = False
    moddat_g_int_CntErr = 0
    
    Do While moddat_g_int_FlgGOK = False
       Screen.MousePointer = 11
       
       g_str_Parame = ""
       g_str_Parame = g_str_Parame & " UPDATE cre_desprocab SET "
       g_str_Parame = g_str_Parame & " descab_cmntes = '" & txt_Comentario.Text & "' "
             
       g_str_Parame = g_str_Parame & " WHERE "
       g_str_Parame = g_str_Parame & " DESCAB_NUMOPE = '" & moddat_g_str_NumOpe & "' and "
       g_str_Parame = g_str_Parame & " DESCAB_FECREG = '" & moddat_g_str_FecRec & "' and "
       g_str_Parame = g_str_Parame & " DESCAB_HORREG = '" & moddat_g_str_FecHip & "' "
       
       If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
          moddat_g_int_CntErr = moddat_g_int_CntErr + 1
          Else
          moddat_g_int_FlgGOK = True
       End If
        
       If moddat_g_int_CntErr = 6 Then
          If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
             Screen.MousePointer = 0
             Exit Sub
             Else
             moddat_g_int_CntErr = 0
          End If
       End If
       Screen.MousePointer = 0
    Loop
          
    r_bol_Estado = fs_guardar_ctaPromotor
    If (r_bol_Estado = False) Then
        Exit Sub
    End If
    
    Screen.MousePointer = 0
    If (moddat_g_int_CntErr = 0) Then
        MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
        frm_RegDes_01.fs_Buscar_Creditos
        Unload Me
    End If
End Sub

Private Function fs_guardar_ctaPromotor() As Boolean
Dim r_int_Fila As Integer

      'GUARDAR CUENTAS BANCARIAS
      fs_guardar_ctaPromotor = True
      
      If (Len(Trim(moddat_g_str_NumOpe)) > 0 And Len(Trim(moddat_g_str_FecRec)) > 0 And Len(Trim(moddat_g_str_FecHip)) > 0) Then
          For r_int_Fila = 1 To grd_Listad_Dsm.Rows - 1
              If (UCase(Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 12))) = Trim("I")) Then
              
                  g_str_Parame = ""
                  g_str_Parame = "INSERT INTO CRE_DESPRODAT ("
                  g_str_Parame = g_str_Parame & "DESDAT_NUMOPE, "
                  g_str_Parame = g_str_Parame & "DESDAT_FECREG, "
                  g_str_Parame = g_str_Parame & "DESDAT_HORREG, "
                  g_str_Parame = g_str_Parame & "DESDAT_CODBCO, "
                  g_str_Parame = g_str_Parame & "DESDAT_NUMCTA, "
                  g_str_Parame = g_str_Parame & "DESDAT_FRMDES, "
                  g_str_Parame = g_str_Parame & "DESDAT_NUMDES, "
                  g_str_Parame = g_str_Parame & "DESDAT_FCHDES, "
                  g_str_Parame = g_str_Parame & "DESDAT_IMPORT, "
                  g_str_Parame = g_str_Parame & "DESDAT_ANOMBR, "
                  g_str_Parame = g_str_Parame & "DESDAT_DESCRI, "
                  g_str_Parame = g_str_Parame & "DESDAT_NUMITE, "
                  g_str_Parame = g_str_Parame & "DESDAT_TIPMTO, "
                  g_str_Parame = g_str_Parame & "SEGUSUCRE, "
                  g_str_Parame = g_str_Parame & "SEGFECCRE, "
                  g_str_Parame = g_str_Parame & "SEGHORCRE, "
                  g_str_Parame = g_str_Parame & "SEGPLTCRE, "
                  g_str_Parame = g_str_Parame & "SEGTERCRE, "
                  g_str_Parame = g_str_Parame & "SEGSUCCRE, "
                  g_str_Parame = g_str_Parame & "SEGUSUACT, "
                  g_str_Parame = g_str_Parame & "SEGFECACT, "
                  g_str_Parame = g_str_Parame & "SEGHORACT, "
                  g_str_Parame = g_str_Parame & "SEGPLTACT, "
                  g_str_Parame = g_str_Parame & "SEGTERACT, "
                  g_str_Parame = g_str_Parame & "SEGSUCACT) "
                  g_str_Parame = g_str_Parame & "VALUES ( "
                  g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
                  g_str_Parame = g_str_Parame & moddat_g_str_FecRec & ", "
                  g_str_Parame = g_str_Parame & moddat_g_str_FecHip & ", "
                  g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 5)) & "', " 'DESDAT_CODBCO
                  g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 7)) & "', " 'DESDAT_NUMCTA
                  g_str_Parame = g_str_Parame & Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 0)) & ", " 'DESDAT_FRMDES
                  'g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 9)) & "', " 'DESDAT_NUMDES
                  'g_str_Parame = g_str_Parame & "'" & Format(Trim(grd_Listad_Dsm.TextMatrix(r_int_fila, 10)), "yyyymmdd") & "', " 'DESDAT_FCHDES
                  g_str_Parame = g_str_Parame & "'" & Trim(txt_NroDsm_Dsm.Text) & "', " 'DESDAT_NUMDES
                  g_str_Parame = g_str_Parame & "'" & Format(ipp_FecDsm_Dsm.Text, "yyyymmdd") & "', " 'DESDAT_FCHDES
                  
                  g_str_Parame = g_str_Parame & Format(Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 4)), "########0.00") & ", " 'DESDAT_IMPORT
                  g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 8)) & "', " 'DESDAT_ANOMBR
                  g_str_Parame = g_str_Parame & "'" & Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 11)) & "', " 'DESDAT_DESCRI
                  g_str_Parame = g_str_Parame & grd_Listad_Dsm.TextMatrix(r_int_Fila, 13) & ", " 'DESDAT_NUMITE
                  g_str_Parame = g_str_Parame & grd_Listad_Dsm.TextMatrix(r_int_Fila, 2) & ", " 'DESDAT_TIPMTO
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
                  g_str_Parame = g_str_Parame & "'" & Format(date, "YYYYMMDD") & "', "
                  g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
                  g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
                  g_str_Parame = g_str_Parame & "'" & Format(date, "YYYYMMDD") & "', "
                  g_str_Parame = g_str_Parame & "'" & Format(Time, "HHMMSS") & "', "
                  g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
                  g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                  
                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                     moddat_g_int_FlgGOK = False
                  End If
              ElseIf (UCase(Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 12))) = Trim("U") Or _
                      UCase(Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 12))) = Trim("S")) Then
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & "UPDATE CRE_DESPRODAT SET "
                  g_str_Parame = g_str_Parame & "       DESDAT_FRMDES =  " & grd_Listad_Dsm.TextMatrix(r_int_Fila, 0) & ","
                  'g_str_Parame = g_str_Parame & "       DESDAT_NUMDES = '" & grd_Listad_Dsm.TextMatrix(r_int_fila, 9) & "',"
                  'g_str_Parame = g_str_Parame & "       DESDAT_FCHDES = '" & Format(grd_Listad_Dsm.TextMatrix(r_int_fila, 10), "yyyymmdd") & "',"
                  g_str_Parame = g_str_Parame & "       DESDAT_NUMDES = '" & Trim(txt_NroDsm_Dsm.Text) & "',"
                  g_str_Parame = g_str_Parame & "       DESDAT_FCHDES = '" & Format(ipp_FecDsm_Dsm.Text, "yyyymmdd") & "',"
                  
                  g_str_Parame = g_str_Parame & "       DESDAT_IMPORT =  " & Format(grd_Listad_Dsm.TextMatrix(r_int_Fila, 4), "########0.00") & ","
                  g_str_Parame = g_str_Parame & "       DESDAT_DESCRI = '" & grd_Listad_Dsm.TextMatrix(r_int_Fila, 11) & "',"
                  g_str_Parame = g_str_Parame & "       DESDAT_ANOMBR = '" & grd_Listad_Dsm.TextMatrix(r_int_Fila, 8) & "',"
                  g_str_Parame = g_str_Parame & "       DESDAT_CODBCO = '" & Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 5)) & "', "
                  g_str_Parame = g_str_Parame & "       DESDAT_NUMCTA = '" & Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 7)) & "', "
                  g_str_Parame = g_str_Parame & "       DESDAT_TIPMTO =  " & grd_Listad_Dsm.TextMatrix(r_int_Fila, 2) & ","
                  g_str_Parame = g_str_Parame & "       SEGUSUACT='" & modgen_g_str_CodUsu & "',"
                  g_str_Parame = g_str_Parame & "       SEGFECACT='" & Format(date, "YYYYMMDD") & "',"
                  g_str_Parame = g_str_Parame & "       SEGHORACT='" & Format(Time, "HHMMSS") & "',"
                  g_str_Parame = g_str_Parame & "       SEGPLTACT='" & UCase(App.EXEName) & "',"
                  g_str_Parame = g_str_Parame & "       SEGTERACT='" & modgen_g_str_NombPC & "',"
                  g_str_Parame = g_str_Parame & "       SEGSUCACT='" & modgen_g_str_CodSuc & "' "
                  g_str_Parame = g_str_Parame & " WHERE TRIM(DESDAT_NUMOPE) ='" & Trim(moddat_g_str_NumOpe) & "' "
                  g_str_Parame = g_str_Parame & "   AND DESDAT_FECREG = " & moddat_g_str_FecRec
                  g_str_Parame = g_str_Parame & "   AND DESDAT_HORREG = " & moddat_g_str_FecHip
                  g_str_Parame = g_str_Parame & "   AND DESDAT_NUMITE = " & grd_Listad_Dsm.TextMatrix(r_int_Fila, 13)

                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                     moddat_g_int_FlgGOK = False
                  End If
              ElseIf (UCase(Trim(grd_Listad_Dsm.TextMatrix(r_int_Fila, 12))) = Trim("D")) Then
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & "DELETE FROM CRE_DESPRODAT "
                  g_str_Parame = g_str_Parame & " WHERE DESDAT_NUMOPE = '" & Trim(moddat_g_str_NumOpe) & "' "
                  g_str_Parame = g_str_Parame & "   AND DESDAT_FECREG =  " & moddat_g_str_FecRec
                  g_str_Parame = g_str_Parame & "   AND DESDAT_HORREG =  " & moddat_g_str_FecHip
                  g_str_Parame = g_str_Parame & "   AND DESDAT_NUMITE =  " & grd_Listad_Dsm.TextMatrix(r_int_Fila, 13)
                  
                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
                     moddat_g_int_FlgGOK = False
                  End If
              End If
              
              If moddat_g_int_FlgGOK = False Then
                 Screen.MousePointer = 0
                 MsgBox "No se pudo completar la grabación de los datos.", vbInformation, modgen_g_str_NomPlt
                 fs_guardar_ctaPromotor = False
                 Exit Function
              End If
          Next
      End If
End Function

Private Sub cmd_Rechazar_Click()
Dim r_bol_Estado  As Boolean

   If Len(Trim(moddat_g_str_NumOpe)) = 0 Then
      MsgBox "Tiene que Haber un Nro. Operación.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Validando
   'If (CDbl(pnl_SumTot_Dsm.Caption) <> 0) Then
   '    If (CDbl(pnl_SumTot_Dsm.Caption) <> CDbl(CStr(l_dbl_ImpPtm))) Then
   '        SSTab1.Tab = 6
   '        MsgBox "El préstamo total no es igual al distribuido en la pestaña datos del desembolso", vbExclamation, modgen_g_str_NomPlt
   '        Exit Sub
   '    End If
   'End If
   
   Screen.MousePointer = 11
   'descab_fecreg , descab_horreg
   g_str_Parame = "select DESCAB_NUMOPE, DESCAB_CODEST from CRE_DESPROCAB where DESCAB_NUMOPE = '" & moddat_g_str_NumOpe & "'   "
   g_str_Parame = g_str_Parame & " AND descab_fecreg = '" & moddat_g_str_FecRec & "'    "
   g_str_Parame = g_str_Parame & " AND descab_horreg = '" & moddat_g_str_FecHip & "'    "
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   If (Trim(CStr(g_rst_Princi!DESCAB_CODEST)) <> Trim(moddat_g_str_CodIte)) Then
      MsgBox "Este registro ya ha cambiado de estado.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   'Actualizando Registro
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
   
      g_str_Parame = "usp_Actualiza_cre_desprocab("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecRec & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_FecHip & "', "
      g_str_Parame = g_str_Parame & "'3', " 'rechazar
      g_str_Parame = g_str_Parame & "'7', "
      
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_FESLNT
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLEG
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_FEENNT
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNLE2
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_FERELG
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNOPE
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_FERECE
      g_str_Parame = g_str_Parame & "'', " 'DESCAB_CMNOP2
      g_str_Parame = g_str_Parame & "'" & txt_Comentario.Text & "', "
      '----------------------------
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      Screen.MousePointer = 0
   Loop
   
   r_bol_Estado = fs_guardar_ctaPromotor
   If (r_bol_Estado = False) Then
       Exit Sub
   End If
   
   'Enviando Correo Electrónico
   Call fs_Envia_Correo("RECHAZO")
   
   'Imprime liquidacion
   Screen.MousePointer = 0
   If (moddat_g_int_CntErr = 0) Then
       MsgBox "El proceso se grabó exitosamente.", vbInformation, modgen_g_str_NomPlt
       frm_RegDes_01.fs_Buscar_Creditos
       Unload Me
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom
   
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_Produc.Caption = Trim(moddat_g_str_NomPrd)
   pnl_EstadoActual.Caption = moddat_g_str_Situac
   pnl_NomCli.Caption = Trim(CStr(moddat_g_int_TipDoc)) & "-" & Trim(moddat_g_str_NumDoc) & " / " & Trim(moddat_g_str_NomCli)

   Call fs_Inicia
   Call moddat_gf_Cargar_AgrPrd
   
   'Buscando información de la solicitud
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)
   Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   Call modmip_gs_DatInm(grd_Listad(1), True)
   l_str_DocPrm = ""
     
   Call fs_DatInm_Aux
   Call fs_DatLeg
   Call fs_DatDes
   Call fs_CalcMto
   Call modmip_gs_DatCre(grd_Listad(4), r_arr_Mtz)
   Call fs_Dat_Evaluacion
   Call fs_PryCta
   
   SSTab1.Tab = 6
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_PryCta()
Dim r_rst_Princi  As ADODB.Recordset
Dim r_str_Cadena  As String
    
    Call moddat_gs_Carga_LisIte_Combo(cmb_FrmDsm_Dsm, 1, "376")
    Call moddat_gs_Carga_LisIte_Combo(cmb_TipMto_Dsm, 1, "132")
    
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " SELECT DISTINCT A.Ctaban_Codbco, TRIM(B.PARDES_DESCRI) AS NOM_BANCO  "
    g_str_Parame = g_str_Parame & "   FROM PRY_CTABAN A  "
    g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 513 AND B.PARDES_CODITE = A.Ctaban_Codbco  "
    g_str_Parame = g_str_Parame & "  WHERE A.CTABAN_CODPRY = '" & CStr(pnl_Prycto_Dsm.Tag) & "'"
    g_str_Parame = g_str_Parame & "    AND A.CTABAN_TIPMON = " & CStr(l_str_CodMod)
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
   
    ReDim l_arr_Bancos(0)
    cmb_EntFin_Dsm.Clear
    
    If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
       g_rst_Princi.MoveFirst
       Do While Not g_rst_Princi.EOF
          ReDim Preserve l_arr_Bancos(UBound(l_arr_Bancos) + 1)
          l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Codigo = Trim$(g_rst_Princi!Ctaban_Codbco)
          l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Nombre = Trim$(g_rst_Princi!NOM_BANCO & "")
          cmb_EntFin_Dsm.AddItem Trim$(g_rst_Princi!NOM_BANCO & "")
          
          g_rst_Princi.MoveNext
       Loop
    End If

    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
    '-------------------------------------------------------
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT CT.* "
    g_str_Parame = g_str_Parame & "  FROM PRY_CTABAN CT "
    g_str_Parame = g_str_Parame & " WHERE CT.CTABAN_CODPRY = '" & CStr(pnl_Prycto_Dsm.Tag) & "'"
    g_str_Parame = g_str_Parame & "   AND CT.CTABAN_TIPMON = " & CStr(l_str_CodMod)
    'g_str_Parame = g_str_Parame & "   AND CT.CTABAN_SITUAC = 1 "
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
   
    ReDim l_arr_CtaBco(0)
    
    If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
       g_rst_Princi.MoveFirst
       Do While Not g_rst_Princi.EOF
          ReDim Preserve l_arr_CtaBco(UBound(l_arr_CtaBco) + 1)
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_Codigo = Trim$(g_rst_Princi!Ctaban_Codbco)
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_Nombre = Trim$(g_rst_Princi!CtaBan_NumCta)
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_FlgAso = Trim$(g_rst_Princi!ctaban_Situac)
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_Refere = Trim$(g_rst_Princi!CTABAN_NUMCCI & "")
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_NomCli = Trim$(g_rst_Princi!CTABAN_ANOMDE & "")
          l_arr_CtaBco(UBound(l_arr_CtaBco)).Genera_ConHip = Trim$(g_rst_Princi!CTABAN_NOMCHQ & "")
          
          g_rst_Princi.MoveNext
       Loop
    End If
    
    Call cmd_Dsm_Cancel_Click
    Call fs_sumarDesemPrmt

    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
    
    '******************
    'Detalle de cuentas
    '******************
    If (Trim(moddat_g_str_FecHip) <> "" And Trim(moddat_g_str_FecRec) <> "") Then
        g_str_Parame = ""
        g_str_Parame = g_str_Parame & " SELECT DT.DESDAT_NUMOPE, DT.DESDAT_FECREG, DT.DESDAT_HORREG, DT.DESDAT_CODBCO,  "
        g_str_Parame = g_str_Parame & "        (SELECT A.PARDES_DESCRI FROM MNT_PARDES A WHERE A.PARDES_CODGRP = 513 AND A.PARDES_CODITE = DT.DESDAT_CODBCO) AS BANCO,  "
        g_str_Parame = g_str_Parame & "        DT.DESDAT_NUMCTA, DT.DESDAT_FRMDES, DT.DESDAT_NUMITE, DT.DESDAT_TIPMTO,  "
        g_str_Parame = g_str_Parame & "        (SELECT A.PARDES_DESCRI FROM MNT_PARDES A WHERE A.PARDES_CODGRP = 376 AND A.PARDES_CODITE = DT.DESDAT_FRMDES) AS TIPODESEMBOLSO,  "
        g_str_Parame = g_str_Parame & "        (SELECT A.PARDES_DESCRI FROM MNT_PARDES A WHERE A.PARDES_CODGRP = 132 AND A.PARDES_CODITE = DT.DESDAT_TIPMTO) AS TIPOMONTO,  "
        g_str_Parame = g_str_Parame & "        DT.DESDAT_NUMDES, DT.DESDAT_FCHDES, DT.DESDAT_DESCRI, "
        g_str_Parame = g_str_Parame & "        DT.DESDAT_ANOMBR , DT.DESDAT_IMPORT  "
        g_str_Parame = g_str_Parame & "   FROM CRE_DESPRODAT DT "
        g_str_Parame = g_str_Parame & "  WHERE DT.DESDAT_NUMOPE = '" & moddat_g_str_NumOpe & "' "
        g_str_Parame = g_str_Parame & "    AND DT.DESDAT_FECREG = " & moddat_g_str_FecRec
        g_str_Parame = g_str_Parame & "    AND DT.DESDAT_HORREG = " & moddat_g_str_FecHip
   
        If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
           Exit Sub
        End If
               
        If Not (r_rst_Princi.EOF And r_rst_Princi.BOF) Then
           r_rst_Princi.MoveFirst
           Do While Not r_rst_Princi.EOF
              grd_Listad_Dsm.Rows = grd_Listad_Dsm.Rows + 1
              grd_Listad_Dsm.Row = grd_Listad_Dsm.Rows - 1
 
              grd_Listad_Dsm.Col = 0
              grd_Listad_Dsm.Text = Trim(r_rst_Princi!DESDAT_FRMDES)
              
              grd_Listad_Dsm.Col = 1
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!TIPODESEMBOLSO) = True, "", Trim(r_rst_Princi!TIPODESEMBOLSO))
              
              grd_Listad_Dsm.Col = 2
              grd_Listad_Dsm.Text = Trim(r_rst_Princi!DESDAT_TIPMTO)
              
              grd_Listad_Dsm.Col = 3
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!TIPOMONTO) = True, "", Trim(r_rst_Princi!TIPOMONTO))
  
              grd_Listad_Dsm.Col = 4
              grd_Listad_Dsm.Text = gf_FormatoNumero(r_rst_Princi!DESDAT_IMPORT, 12, 2)
              
              grd_Listad_Dsm.Col = 5
              grd_Listad_Dsm.Text = Trim(r_rst_Princi!DESDAT_CODBCO & "")
            
              grd_Listad_Dsm.Col = 6
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!BANCO) = True, "", Trim(r_rst_Princi!BANCO))
              
              grd_Listad_Dsm.Col = 7
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!DESDAT_NUMCTA) = True, "", Trim(r_rst_Princi!DESDAT_NUMCTA))
              
              grd_Listad_Dsm.Col = 8
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!DESDAT_ANOMBR) = True, "", Trim(r_rst_Princi!DESDAT_ANOMBR))
                            
              grd_Listad_Dsm.Col = 9
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!DESDAT_NUMDES) = True, "", Trim(r_rst_Princi!DESDAT_NUMDES))
              txt_NroDsm_Dsm.Text = grd_Listad_Dsm.Text
                            
              If IsNull(r_rst_Princi!DESDAT_FCHDES) = False Then
                 grd_Listad_Dsm.Col = 10
                 grd_Listad_Dsm.Text = gf_FormatoFecha(r_rst_Princi!DESDAT_FCHDES)
                 ipp_FecDsm_Dsm.Text = grd_Listad_Dsm.Text
              End If
              
              grd_Listad_Dsm.Col = 11
              grd_Listad_Dsm.Text = IIf(IsNull(r_rst_Princi!DESDAT_DESCRI) = True, "", Trim(r_rst_Princi!DESDAT_DESCRI))
                        
              grd_Listad_Dsm.Col = 12
              grd_Listad_Dsm.Text = "S"
              
              grd_Listad_Dsm.Col = 13
              grd_Listad_Dsm.Text = Trim(r_rst_Princi!DESDAT_NUMITE)
               
              r_rst_Princi.MoveNext
              DoEvents
           Loop
           Call gs_UbiIniGrid(grd_Listad_Dsm)
        End If
    End If
    Call fs_sumarDesemPrmt
    
    If (Trim(pnl_Prycto_Dsm.Tag) = "") Then
        cmd_Dsm_Nuevo.Enabled = False
        cmd_Dsm_Borrar.Enabled = False
        cmd_Dsm_Editar.Enabled = False
    End If
    If (CInt(moddat_g_int_CodIns) = CInt("000005")) Then
        'LEGAL 2DA PARTE
        If (CInt(moddat_g_str_CodIte) = CInt("000010") Or CInt(moddat_g_str_CodIte) = CInt("000009")) Then
            cmd_Dsm_Nuevo.Enabled = False
            cmd_Dsm_Borrar.Enabled = False
            cmd_Dsm_Editar.Enabled = False
        End If
    End If
    If (cmd_Grabar.Enabled = False) Then
        cmd_Dsm_Nuevo.Enabled = False
        cmd_Dsm_Borrar.Enabled = False
        cmd_Dsm_Editar.Enabled = False
    End If
    If (moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = CInt("000001")) Then
        'OPERACIONES  -  LEGAL
        lbl_NumDsm_Dsm.Visible = False
        txt_NroDsm_Dsm.Visible = False
        lbl_FchDsm_Dsm.Visible = False
        ipp_FecDsm_Dsm.Visible = False
    ElseIf (moddat_g_int_CodIns = CInt("000003") Or moddat_g_int_CodIns = CInt("000004") Or moddat_g_int_CodIns = CInt("000005")) Then
        'TESORERIA  -  OPERACIONES 2DA PARTE   -   LEGAL 2DA PARTE
        lbl_NumDsm_Dsm.Visible = True
        txt_NroDsm_Dsm.Visible = True
        lbl_FchDsm_Dsm.Visible = True
        ipp_FecDsm_Dsm.Visible = True
    End If
    If (moddat_g_int_CodIns = CInt("000004") Or moddat_g_int_CodIns = CInt("000005")) Then ''LEGAL 2DA PARTE
        'OPERACIONES 2DA PARTE    -   LEGAL 2DA PARTE
        cmd_Dsm_Nuevo.Enabled = False
        cmd_Dsm_Borrar.Enabled = False
        cmd_Dsm_Editar.Enabled = False
    End If
End Sub
Private Sub gs_Carga_EntFin()
   cmb_EntFin_Dsm.Clear
   ReDim l_arr_Bancos(0)
      
   ReDim Preserve l_arr_Bancos(UBound(l_arr_Bancos) + 1)
   l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Codigo = Trim$(l_str_CodBan)
   l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Nombre = Trim$(l_str_PryBan)
   l_arr_Bancos(UBound(l_arr_Bancos)).Genera_TipVal = 0
   l_arr_Bancos(UBound(l_arr_Bancos)).Genera_Cantid = 0
End Sub

Public Sub fs_Dat_Evaluacion()
Dim r_str_frmDesem As String
Dim r_str_feslnt As String
Dim r_str_nrodes As String
Dim r_str_ferece As String
Dim r_str_feennt As String
Dim r_str_cmnleg As String
Dim r_str_cmnope As String
Dim r_str_cmntes As String
Dim r_str_cmnop2 As String
Dim r_str_cmnle2 As String

Dim r_str_Legal As String
Dim r_str_Oper As String
Dim r_str_Teso As String
Dim r_str_Oper_2 As String
Dim r_str_Legal_2 As String

Dim r_str_LegFec As String
Dim r_str_OpeFec As String
Dim r_str_TesFec As String
Dim r_str_OpeFec_2 As String
Dim r_str_LegFec_2 As String

g_str_Parame = "  SELECT to_number(PARDES_CODITE) PARDES_CODITE, trim(PARDES_DESCRI) as Instancia,  "
g_str_Parame = g_str_Parame & "  (select det.desdet_fecfin from cre_desprodet det  "
g_str_Parame = g_str_Parame & "  where det.desdet_numope = '" & moddat_g_str_NumOpe & "'  "
g_str_Parame = g_str_Parame & "  and det.desdet_fecreg = '" & moddat_g_str_FecRec & "'  "
g_str_Parame = g_str_Parame & "  and det.desdet_horreg = '" & moddat_g_str_FecHip & "'  "
g_str_Parame = g_str_Parame & "  and par.PARDES_CODITE = det.desdet_codarea) as fechaEnvio  "

g_str_Parame = g_str_Parame & "  FROM MNT_PARDES par WHERE par.PARDES_CODGRP = '374'  "
g_str_Parame = g_str_Parame & "  and par.PARDES_CODITE <> '000000' AND par.PARDES_SITUAC = 1  "
g_str_Parame = g_str_Parame & "  ORDER BY PARDES_CODITE ASC  "

If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
   Exit Sub
End If

If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
   Select Case g_rst_Princi!PARDES_CODITE
          Case 1
                r_str_Legal = g_rst_Princi!INSTANCIA
                If (IsNull(g_rst_Princi!fechaEnvio) = False) Then
                    r_str_LegFec = g_rst_Princi!fechaEnvio
                End If
          Case 2
                r_str_Oper = g_rst_Princi!INSTANCIA
                If (IsNull(g_rst_Princi!fechaEnvio) = False) Then
                    r_str_OpeFec = g_rst_Princi!fechaEnvio
                End If
          Case 3
                r_str_Teso = g_rst_Princi!INSTANCIA
                If (IsNull(g_rst_Princi!fechaEnvio) = False) Then
                    r_str_TesFec = g_rst_Princi!fechaEnvio
                End If
          Case 4
                r_str_Oper_2 = g_rst_Princi!INSTANCIA
                If (IsNull(g_rst_Princi!fechaEnvio) = False) Then
                    r_str_OpeFec_2 = g_rst_Princi!fechaEnvio
                End If
          Case 5
                r_str_Legal_2 = g_rst_Princi!INSTANCIA
                If (IsNull(g_rst_Princi!fechaEnvio) = False) Then
                    r_str_LegFec_2 = g_rst_Princi!fechaEnvio
                End If
   End Select
   g_rst_Princi.MoveNext
   Loop
End If

g_rst_Princi.Close
Set g_rst_Princi = Nothing
'--------------------------------------------------------------------------------------------------------------
   Call gs_LimpiaGrid(grd_Listad(7))
 
   g_str_Parame = "" '-----trae la Cabecera----
   g_str_Parame = g_str_Parame & "SELECT descab_numope, descab_codarea, descab_codest,descab_fecreg, descab_horreg, "
   g_str_Parame = g_str_Parame & "       descab_feslnt, descab_ferece, descab_feennt, descab_cmnleg, descab_cmnope, "
   g_str_Parame = g_str_Parame & "       descab_cmntes , descab_cmnop2, descab_cmnle2, descab_FERELG "
   g_str_Parame = g_str_Parame & "  FROM cre_desprocab cab "
   g_str_Parame = g_str_Parame & " WHERE cab.descab_numope = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND cab.DESCAB_FECREG = '" & moddat_g_str_FecRec & "' "
   g_str_Parame = g_str_Parame & "   AND cab.DESCAB_HORREG = '" & moddat_g_str_FecHip & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      txt_Comentario.Text = IIf(IsNull(g_rst_Princi!descab_cmntes) = True, "", g_rst_Princi!descab_cmntes)
                   
      Do While Not g_rst_Princi.EOF
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = "Instancia"
             grd_Listad(7).Col = 1
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = r_str_Legal
                  
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha de Solicitud Notaria"
             If (IsNull(g_rst_Princi!descab_feslnt) = False) Then
                 grd_Listad(7).Col = 1
                 grd_Listad(7).Text = gf_FormatoFecha(g_rst_Princi!descab_feslnt)
             End If
        
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Comentario Legal"
             grd_Listad(7).Col = 1
             grd_Listad(7).Text = IIf(IsNull(g_rst_Princi!descab_cmnleg), "", g_rst_Princi!descab_cmnleg)
             
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Envío"
             grd_Listad(7).Col = 1
             If (Len(Trim((r_str_LegFec))) <> 0) Then
                 grd_Listad(7).Text = gf_FormatoFecha(r_str_LegFec)
             End If
'-----------------------------------------------------------------------------------------------------
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = "Instancia"
             grd_Listad(7).Col = 1
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = r_str_Oper
         
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Comentario Operaciones"
             grd_Listad(7).Col = 1
             grd_Listad(7).Text = IIf(IsNull(g_rst_Princi!descab_cmnope), "", g_rst_Princi!descab_cmnope)
             
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Envío"
             grd_Listad(7).Col = 1
             If (Len(Trim((r_str_OpeFec))) <> 0) Then
                 grd_Listad(7).Text = gf_FormatoFecha(r_str_OpeFec)
             End If
'-----------------------------------------------------------------------------------------------------
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = "Instancia"
             grd_Listad(7).Col = 1
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = r_str_Teso
                                   
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Comentario Tesoreria"
             grd_Listad(7).Col = 1
             grd_Listad(7).Text = IIf(IsNull(g_rst_Princi!descab_cmntes), "", g_rst_Princi!descab_cmntes)
             
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Envío"
             grd_Listad(7).Col = 1
             If (Len(Trim((r_str_TesFec))) <> 0) Then
                 grd_Listad(7).Text = gf_FormatoFecha(r_str_TesFec)
             End If
'-----------------------------------------------------------------------------------------------------
            grd_Listad(7).Rows = grd_Listad(7).Rows + 1
            grd_Listad(7).Row = grd_Listad(7).Rows - 1
            grd_Listad(7).Col = 0
            grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(7).Text = "Instancia"
            grd_Listad(7).Col = 1
            grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(7).Text = r_str_Oper_2
                  
            grd_Listad(7).Rows = grd_Listad(7).Rows + 1
            grd_Listad(7).Row = grd_Listad(7).Rows - 1
            grd_Listad(7).Col = 0
            grd_Listad(7).Text = "Fecha Recepcion Const. Desembolso"
            If (IsNull(g_rst_Princi!descab_ferece) = False) Then
                grd_Listad(7).Col = 1
                grd_Listad(7).Text = gf_FormatoFecha(g_rst_Princi!descab_ferece)
            End If
         
            grd_Listad(7).Rows = grd_Listad(7).Rows + 1
            grd_Listad(7).Row = grd_Listad(7).Rows - 1
            grd_Listad(7).Col = 0
            grd_Listad(7).Text = "Comentario Operaciones 2da Parte"
            grd_Listad(7).Col = 1
            grd_Listad(7).Text = IIf(IsNull(g_rst_Princi!descab_cmnop2), "", g_rst_Princi!descab_cmnop2)
            
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Envío"
             grd_Listad(7).Col = 1
             If (Len(Trim((r_str_OpeFec_2))) <> 0) Then
                 grd_Listad(7).Text = gf_FormatoFecha(r_str_OpeFec_2)
             End If
'-----------------------------------------------------------------------------------------------------
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = "Instancia"
             grd_Listad(7).Col = 1
             grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
             grd_Listad(7).Text = r_str_Legal_2
         
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha de Recepción"
             If (IsNull(g_rst_Princi!descab_FERELG) = False) Then
                 grd_Listad(7).Col = 1
                 grd_Listad(7).Text = gf_FormatoFecha(g_rst_Princi!descab_FERELG)
             End If
             
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha de Entrega Notaria"
             If (IsNull(g_rst_Princi!descab_feennt) = False) Then
                 grd_Listad(7).Col = 1
                 grd_Listad(7).Text = gf_FormatoFecha(g_rst_Princi!descab_feennt)
             End If
                                                
             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Comentario Legal 2da Parte"
             grd_Listad(7).Col = 1
             grd_Listad(7).Text = IIf(IsNull(g_rst_Princi!descab_cmnle2), "", g_rst_Princi!descab_cmnle2)

             grd_Listad(7).Rows = grd_Listad(7).Rows + 1
             grd_Listad(7).Row = grd_Listad(7).Rows - 1
             grd_Listad(7).Col = 0
             grd_Listad(7).Text = "Fecha Termino"
             grd_Listad(7).Col = 1
             If (Len(Trim((r_str_LegFec_2))) <> 0) Then
                 grd_Listad(7).Text = gf_FormatoFecha(r_str_LegFec_2)
             End If
             
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad(7))
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   Dim r_str_TipCli     As String
   
   r_str_TipCli = ""

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(0).Redraw = False
      
      If p_Indice = 1 Then
         r_str_TipCli = " (Cónyuge)"
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      End If
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Documento de Identidad" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DATGEN_TIPDOC)) & " - " & Trim(g_rst_Princi!DATGEN_NUMDOC & "")
   
      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Apellidos y Nombres" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
      
      If p_Indice = 0 Then
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Estado Civil"
         
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DatGen_RegCyg), "")
         
         If g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
            moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
         End If
      End If

      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Celular" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      
      If p_Indice = 0 Then
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Domicilio"
         
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Referencia"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
      
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Teléfono Domicilio"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      End If
      
      grd_Listad(0).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(0))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatInm_Aux()
Dim r_str_Cadena As String
   l_str_PryBan = ""
   l_str_CodBan = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SL.*, EL.EVALEG_FEENIN, PY.DATGEN_VENTDO, PY.DATGEN_VENNDO, PY.DATGEN_CONTDO, PY.DATGEN_CONNDO, "
   g_str_Parame = g_str_Parame & "       (SELECT TRIM(A.PARDES_DESCRI) FROM MNT_PARDES A WHERE A.PARDES_CODGRP = 513 AND A.PARDES_CODITE = SL.SOLINM_PRYBCO) AS BANCO, "
   g_str_Parame = g_str_Parame & "       (SELECT TRIM(B.DATGEN_TITULO) FROM PRY_DATGEN B WHERE B.DATGEN_CODIGO = SL.SOLINM_PRYCOD) AS PROYECTO "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLINM SL "
   g_str_Parame = g_str_Parame & "   LEFT JOIN PRY_DATGEN PY ON PY.DATGEN_CODIGO = SL.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "   LEFT JOIN TRA_EVALEG EL ON EL.EVALEG_NUMSOL = SL.SOLINM_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE SL.SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_Prycto_Dsm.Caption = IIf(IsNull(g_rst_Princi!PROYECTO) = True, "", g_rst_Princi!PROYECTO)
      pnl_Prycto_Dsm.Tag = IIf(IsNull(g_rst_Princi!SOLINM_PRYCOD) = True, "", g_rst_Princi!SOLINM_PRYCOD)
      If Not IsNull(g_rst_Princi!BANCO) And Trim(CStr(g_rst_Princi!SOLINM_PRYBCO & "")) <> "888888" Then
         l_str_PryBan = Trim(CStr(g_rst_Princi!BANCO))
         l_str_CodBan = Trim(CStr(g_rst_Princi!SOLINM_PRYBCO))
      End If
            
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         'CREDITOS ANTIGUOS
         If (Len(Trim(g_rst_Princi!SOLINM_TIPDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_RAZSOC_PRO)) > 0) Then
             l_str_Prmtor = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO) & _
                            " / " & moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
             l_str_DocPrm = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
         Else
             If (Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0) Then
                 r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!DATGEN_VENTDO, g_rst_Princi!DATGEN_VENNDO)
                 If (Len(Trim(r_str_Cadena)) > 0) Then
                     l_str_Prmtor = CStr(g_rst_Princi!DATGEN_VENTDO) & "-" & Trim(g_rst_Princi!DATGEN_VENNDO) & _
                                    " / " & moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
                     l_str_DocPrm = CStr(g_rst_Princi!DATGEN_VENTDO) & "-" & Trim(g_rst_Princi!DATGEN_VENNDO)
                 End If
             End If
         End If
      Else
      'CREDITOS NUEVOS
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Then
            If (Len(Trim(g_rst_Princi!SOLINM_TIPDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)) > 0 And Len(Trim(g_rst_Princi!SOLINM_RAZSOC_PRO)) > 0) Then
                l_str_Prmtor = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO) & _
                               " / " & Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
                l_str_DocPrm = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            Else
                If (Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0) Then
                    r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!DATGEN_VENTDO, g_rst_Princi!DATGEN_VENNDO)
                    If (Len(Trim(r_str_Cadena)) > 0) Then
                        l_str_Prmtor = CStr(g_rst_Princi!DATGEN_VENTDO) & "-" & Trim(g_rst_Princi!DATGEN_VENNDO) & _
                               " / " & r_str_Cadena
                        l_str_DocPrm = CStr(g_rst_Princi!DATGEN_VENTDO) & "-" & Trim(g_rst_Princi!DATGEN_VENNDO)
                    End If
                End If
            End If
         Else
            '*********BIEN FUTURO**********
            r_str_Cadena = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            If (Len(Trim(r_str_Cadena)) > 0) Then
                l_str_Prmtor = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO) & _
                               " / " & r_str_Cadena
                l_str_DocPrm = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            End If
         End If
      End If
      
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatLeg()
   Call gs_LimpiaGrid(grd_Listad(6))
   l_int_MonCvt = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVALEG "
   g_str_Parame = g_str_Parame & " WHERE EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      txt_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLG1 & "") & Trim(g_rst_Princi!EVALEG_INFLG2 & "") & Trim(g_rst_Princi!EVALEG_INFLG3 & "") & Trim(g_rst_Princi!EVALEG_INFLG4 & "")
      txt_ComCre.Text = "Fecha de Comité de Créditos: " & gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCOM)) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Trim(g_rst_Princi!EVALEG_OBSCOM & "")
      
      If g_rst_Princi!EVALEG_FECCVT > 0 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Fecha Firma Contrato Compra Venta"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCVT))
         
         If Not IsNull(g_rst_Princi!EVALEG_TCASBS) Then
            If g_rst_Princi!EVALEG_TCASBS > 0 Then
               grd_Listad(6).Rows = grd_Listad(6).Rows + 1
               grd_Listad(6).Row = grd_Listad(6).Rows - 1
               grd_Listad(6).Col = 0
               grd_Listad(6).Text = "Tipo de Cambio SBS"
               
               grd_Listad(6).Col = 1
               grd_Listad(6).Text = Format(g_rst_Princi!EVALEG_TCASBS, "###,##0.0000")
            End If
         End If
      
         If g_rst_Princi!EVALEG_TCACVT > 0 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Tipo de Cambio aplicado"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = Format(g_rst_Princi!EVALEG_TCACVT, "###,##0.0000")
         End If
      End If
      
      If Not IsNull(g_rst_Princi!EVALEG_MONCVT) Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Moneda Compra-Venta"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!EVALEG_MONCVT)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Valor Compra-Venta"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONCVT) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_COMVTA, 12, 2)
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Aporte Propio"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONCVT) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_APOPRO, 12, 2)
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Monto Préstamo"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONCVT) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_MTOPRE, 12, 2)
      End If
      
      If grd_Listad(6).Rows = 0 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      Else
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
      End If
      
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Fecha Firma Contrato (Crédito)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Notaria"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("509", g_rst_Princi!EVALEG_CODNOT & "")
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Representante Legal 1"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG1 & "")
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Representante Legal 2"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG2 & "")
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Monto Hipoteca "
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONHIP) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_MTOHIP, 12, 2)
      
      If g_rst_Princi!EVALEG_FECBLQ_INM > 0 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Bloqueo Registral Inscrito"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = "SI"
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Sede Registral"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("511", CStr(g_rst_Princi!EVALEG_SEDREG & ""))
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Fecha Bloqueo (Inmueble)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_INM))
                  
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Doc. Registral (Inmueble)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_INM)
                  
         Select Case g_rst_Princi!EVALEG_TIPDOC_INM
            Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_INM & "")
            Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_INM & "")
            Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_INM & "") & ")"
         End Select
         
         If g_rst_Princi!EVALEG_FLGEST_ES1 = 1 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 2
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Fecha Bloqueo (Estac. 1)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES1))
                       
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Doc. Registral (Estac. 1)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES1)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES1
               Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES1 & "")
               Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES1 & "")
               Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES1 & "") & ")"
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_ES2 = 1 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 2
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Fecha Bloqueo (Estac. 2)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES2))
                        
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Doc. Registral (Estac. 2)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES2)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES2
               Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES2 & "")
               Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES2 & "")
               Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES2 & "") & ")"
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_DEP = 1 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 2
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Fecha Bloqueo (Depósito)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_DEP))
                        
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Doc. Registral (Depósito)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_DEP)
                        
            Select Case g_rst_Princi!EVALEG_TIPDOC_DEP
               Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_DEP & "")
               Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_DEP & "")
               Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_DEP & "") & ")"
            End Select
         End If
      End If
      
      If Not IsNull(g_rst_Princi!EVALEG_MONCVT) Then
         l_int_MonCvt = g_rst_Princi!EVALEG_MONCVT
      End If
      
      Call gs_UbiIniGrid(grd_Listad(6))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatDes()
   Call gs_LimpiaGrid(grd_Listad(3))
   txt_ObsDes.Text = ""
   l_int_FlgCVt = 0

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPDES "
   g_str_Parame = g_str_Parame & " WHERE HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Fecha de Desembolso"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = gf_FormatoFecha(g_rst_Princi!HIPDES_FECDES)
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Tipo de Desembolso"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).Text = "CONTRA " & moddat_gf_Consulta_ParDes("241", g_rst_Princi!HIPDES_TIPGAR)
      
      If g_rst_Princi!HIPDES_TIPGAR = 2 Or g_rst_Princi!HIPDES_TIPGAR = 4 Or g_rst_Princi!HIPDES_TIPGAR = 5 Or g_rst_Princi!HIPDES_TIPGAR = 3 Then
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Forma de Desembolso"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("226", g_rst_Princi!HIPDES_TIPDES)
      End If
      
      If g_rst_Princi!HIPDES_TIPDES = 1 Then
         If Len(Trim(g_rst_Princi!HIPDES_CHECGO & "")) > 0 Then
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. de Cheque"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = Trim(g_rst_Princi!HIPDES_CHECGO & "")
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Banco Emisor (Cuenta)"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("516", g_rst_Princi!HIPDES_BANCGO & "") & " (" & Trim(g_rst_Princi!HIPDES_CTACGO & "") & ")"
         Else
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. de Cheque"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = "CHEQUE NO EMITIDO"
            
            l_int_ChqReg = 1
         End If
      End If
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Importe Desembolsado"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).CellFontName = "Lucida Console"
      grd_Listad(3).CellFontSize = 8
      grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPDES_DESMPR, 12, 2)
      
      If g_rst_Princi!HIPDES_TIPGAR = 4 Then
         If Len(Trim(g_rst_Princi!HIPDES_NUMFIA & "")) > 0 Then
            grd_Listad(3).Rows = grd_Listad(3).Rows + 2
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. Carta Fianza"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = Trim(g_rst_Princi!HIPDES_NUMFIA & "")
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Banco Emisor "
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANFIA)
         
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Fecha Emisión"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPDES_EMIFIA))
         
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Fecha Vencimiento"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPDES_VCTFIA))
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Importe Carta Fianza"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).CellFontName = "Lucida Console"
            grd_Listad(3).CellFontSize = 8
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!HIPDES_MONFIA) & " " & gf_FormatoNumero(g_rst_Princi!HIPDES_IMPFIA, 12, 2)
         Else
            grd_Listad(3).Rows = grd_Listad(3).Rows + 2
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. Carta Fianza"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = "CARTA FIANZA NO RECIBIDA"
            
            l_int_FiaReg = 1
         End If
      End If
      
      If g_rst_Princi!HIPDES_TIPGAR = 5 Then
         If Len(Trim(g_rst_Princi!HIPDES_DOCGAR & "")) > 0 Then
            grd_Listad(3).Rows = grd_Listad(3).Rows + 2
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. Certificado de Participación"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = Trim(g_rst_Princi!HIPDES_DOCGAR & "")
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Banco Emisor "
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BCOGAR)
            
            grd_Listad(3).Rows = grd_Listad(3).Rows + 1
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Importe Certificado"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).CellFontName = "Lucida Console"
            grd_Listad(3).CellFontSize = 8
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!HIPDES_MONGAR) & " " & gf_FormatoNumero(g_rst_Princi!HIPDES_MTOGAR, 12, 2)
         Else
            grd_Listad(3).Rows = grd_Listad(3).Rows + 2
            grd_Listad(3).Row = grd_Listad(3).Rows - 1
            grd_Listad(3).Col = 0
            grd_Listad(3).Text = "Nro. Certificado de Participación"
            
            grd_Listad(3).Col = 1
            grd_Listad(3).Text = "CERTIFICADO NO RECIBIDO"
            
            l_int_CerReg = 1
         End If
      End If
            
      Call gs_UbiIniGrid(grd_Listad(3))
      txt_ObsDes.Text = Trim(g_rst_Princi!HIPDES_OBSERV & "")
      
      If Not IsNull(g_rst_Princi!HIPDES_MONCVT) Then
         l_int_FlgCVt = 1
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_CalcMto()
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_MTOPRE,SOLMAE_FMVBBP, SOLMAE_PBPMTO, SOLMAE_AFPMTO, SOLMAE_BMSMTO, HIPMAE_PRYMCS,  "
   g_str_Parame = g_str_Parame & "       HIPMAE_CVTSOL, HIPMAE_APOSOL, HIPMAE_CVTDOL, HIPMAE_APODOL, HIPMAE_FECESC, HIPMAE_PLAANO, "
   g_str_Parame = g_str_Parame & "       HIPMAE_TASINT, HIPMAE_NUMCUO, HIPMAE_PERGRA, HIPMAE_SEGPRE, HIPMAE_TIPSEG, HIPMAE_CONHIP, SOLMAE_MTOGCI "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE B ON SOLMAE_NUMERO = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 7 OR HIPMAE_SITUAC = 9)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'Datos_Promotor
   pnl_Moneda_Dsm.Caption = moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon))
   pnl_Moneda_Dsm.Tag = moddat_g_int_TipMon
   l_str_CodMod = moddat_g_int_TipMon
   l_str_Moneda = pnl_Moneda_Dsm.Caption
   lbl_Bono_Dsm.Caption = ".."
   l_dbl_ImpPtm = CDbl(g_rst_Princi!HIPMAE_MTOPRE)
   
   If moddat_g_int_TipMon = 1 Then
      If moddat_g_str_CodPrd = "024" Then
         If g_rst_Princi!HIPMAE_PRYMCS = 1 Then
            'VINCULADO
            l_dbl_ImpPtm = CDbl(g_rst_Princi!HIPMAE_MTOPRE) + CDbl(g_rst_Princi!SOLMAE_FMVBBP) + CDbl(g_rst_Princi!SOLMAE_PBPMTO) + CDbl(g_rst_Princi!SOLMAE_BMSMTO) + CDbl(g_rst_Princi!SOLMAE_AFPMTO)
            lbl_Bono_Dsm.Caption = "(INCLUYE BONOS " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP + g_rst_Princi!SOLMAE_PBPMTO + g_rst_Princi!SOLMAE_BMSMTO, "##,###,##0.00") & ")"
         Else
            'NO VINCULADO
            l_dbl_ImpPtm = CDbl(g_rst_Princi!HIPMAE_MTOPRE) + CDbl(g_rst_Princi!SOLMAE_BMSMTO) + CDbl(g_rst_Princi!SOLMAE_AFPMTO)
         End If
      ElseIf InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
         l_dbl_ImpPtm = CDbl(g_rst_Princi!HIPMAE_MTOPRE) + CDbl(g_rst_Princi!SOLMAE_FMVBBP) + CDbl(g_rst_Princi!SOLMAE_AFPMTO) + CDbl(g_rst_Princi!SOLMAE_PBPMTO) + CDbl(g_rst_Princi!SOLMAE_BMSMTO)
         lbl_Bono_Dsm.Caption = "(INCLUYE BONOS " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP + g_rst_Princi!SOLMAE_PBPMTO, "##,###,##0.00") & ")"
      Else
         If moddat_g_str_CodPrd = "011" Then
            l_dbl_ImpPtm = CDbl(g_rst_Princi!HIPMAE_MTOPRE) + CDbl(g_rst_Princi!SOLMAE_AFPMTO)
         End If
      End If
   End If
   l_dbl_ImpPtm = l_dbl_ImpPtm - CDbl(g_rst_Princi!SOLMAE_MTOGCI)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Inicia()
Dim r_str_Parame     As String
Dim r_rst_Genera     As ADODB.Recordset

   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT A.PERMES_CODANO, A.PERMES_CODMES "
   r_str_Parame = r_str_Parame & "  FROM CTB_PERMES A "
   r_str_Parame = r_str_Parame & " WHERE PERMES_CODEMP = '000001' "
   r_str_Parame = r_str_Parame & "   AND PERMES_TIPPER = 1 "
   r_str_Parame = r_str_Parame & "   AND PERMES_SITUAC = 1 "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
      r_rst_Genera.MoveFirst
      modctb_int_PerMes = r_rst_Genera!PERMES_CODMES
      modctb_int_PerAno = r_rst_Genera!PERMES_CODANO
      
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
   End If
   '-----------------------------------------------------------
   If (moddat_g_int_TipRep = 1) Then
       cmd_Aprobar.Enabled = True
       cmd_Rechazar.Enabled = True
       pnl_Titulo.Caption = "Créditos Hipotecarios - Evaluación de Tesoreria"
   Else
       cmd_Aprobar.Enabled = False
       cmd_Rechazar.Enabled = False
       cmd_Grabar.Enabled = False
       pnl_Titulo.Caption = "Créditos Hipotecarios - Consulta de Tesoreria"
       txt_NroDsm_Dsm.Enabled = False
       ipp_FecDsm_Dsm.Enabled = False
       txt_Comentario.Enabled = False
   End If

   'Datos del Cliente
   grd_Listad(0).ColWidth(0) = 3060:   grd_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(0).ColWidth(1) = 7940:   grd_Listad(0).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(0))

   'Datos del Inmueble
   grd_Listad(1).ColWidth(0) = 3060:   grd_Listad(1).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(1).ColWidth(1) = 7940:   grd_Listad(1).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(1))
   
   'Datos Legal
   grd_Listad(6).ColWidth(0) = 3060:   grd_Listad(6).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(6).ColWidth(1) = 7940:   grd_Listad(6).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(6))

   'Datos del Crédito
   grd_Listad(4).ColWidth(0) = 3060:   grd_Listad(4).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(4).ColWidth(1) = 7940:   grd_Listad(4).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(4))

   'Datos del Desembolso
   grd_Listad(3).ColWidth(0) = 3060:   grd_Listad(3).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(3).ColWidth(1) = 7940:   grd_Listad(3).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(3))
         
   'Datos de la Evaluacion
   grd_Listad(7).ColWidth(0) = 3000:   grd_Listad(7).ColAlignment(1) = flexAlignLeftCenter
   grd_Listad(7).ColWidth(1) = 8000:   grd_Listad(7).ColAlignment(1) = flexAlignLeftCenter

   Call gs_LimpiaGrid(grd_Listad(7))
   
  'Datos Desembolso Promotor
   grd_Listad_Dsm.TextMatrix(0, 0) = "ID_FormaPago"
   grd_Listad_Dsm.TextMatrix(0, 1) = "Forma Pago"
   grd_Listad_Dsm.TextMatrix(0, 2) = "ID_TipoMonto"
   grd_Listad_Dsm.TextMatrix(0, 3) = "Tipo Monto"
   grd_Listad_Dsm.TextMatrix(0, 4) = "Importe"
   grd_Listad_Dsm.TextMatrix(0, 5) = "ID_BANCO"
   grd_Listad_Dsm.TextMatrix(0, 6) = "Entidad Financiera"
   grd_Listad_Dsm.TextMatrix(0, 7) = "Nro Cuenta"
   grd_Listad_Dsm.TextMatrix(0, 8) = "A Nombre de"
   grd_Listad_Dsm.TextMatrix(0, 9) = "Nro Desembolso"
   grd_Listad_Dsm.TextMatrix(0, 10) = "Fecha Reg."
   grd_Listad_Dsm.TextMatrix(0, 11) = "Descripcion"
   grd_Listad_Dsm.TextMatrix(0, 12) = "Flag"
   grd_Listad_Dsm.TextMatrix(0, 13) = "NumItem"
   
   If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
       'Legal 1 y Operaciones 1
       grd_Listad_Dsm.TextMatrix(0, 9) = ""
       grd_Listad_Dsm.TextMatrix(0, 10) = ""
   End If
   
   grd_Listad_Dsm.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad_Dsm.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_Dsm.ColAlignment(4) = flexAlignRightCenter
   grd_Listad_Dsm.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad_Dsm.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad_Dsm.ColAlignment(8) = flexAlignLeftCenter
   grd_Listad_Dsm.ColAlignment(9) = flexAlignLeftCenter
   grd_Listad_Dsm.ColAlignment(10) = flexAlignLeftCenter
   grd_Listad_Dsm.ColAlignment(11) = flexAlignLeftCenter
   
   grd_Listad_Dsm.ColWidth(0) = 0      'Id-FormaPago
   grd_Listad_Dsm.ColWidth(1) = 1500   'FormaPago
   grd_Listad_Dsm.ColWidth(2) = 0      'ID_TipoMonto
   grd_Listad_Dsm.ColWidth(3) = 1500   'TipoMonto
   grd_Listad_Dsm.ColWidth(4) = 1200   'Importe
   grd_Listad_Dsm.ColWidth(5) = 0      'ID_Banco
   grd_Listad_Dsm.ColWidth(6) = 2500   'Nom_Banco
   grd_Listad_Dsm.ColWidth(7) = 1800   'Nro_Cuenta
   grd_Listad_Dsm.ColWidth(8) = 3200   'A_Nombre_DE
   grd_Listad_Dsm.ColWidth(9) = 0      'Nro_Desembolso
   grd_Listad_Dsm.ColWidth(10) = 0     'Fec_Desembolso
   grd_Listad_Dsm.ColWidth(11) = 3500  'Descripcion
   grd_Listad_Dsm.ColWidth(12) = 0     'Flag
   grd_Listad_Dsm.ColWidth(13) = 0     'NumItem
   
   If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
       'Legal 1 y Operaciones 1
       'grd_Listad_Dsm.ColWidth(9) = 0 'Nro Desembolso
       'grd_Listad_Dsm.ColWidth(10) = 0 'Fecha Reg.
   Else
       'grd_Listad_Dsm.ColWidth(9) = 1900 'Nro Desembolso
       'grd_Listad_Dsm.ColWidth(10) = 1020  'Fecha Reg.
   End If
      
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 1
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 3
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 4
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 6
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 7
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 8
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   'grd_Listad_Dsm.Row = 0
   'grd_Listad_Dsm.Col = 9
   'grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   'grd_Listad_Dsm.CellBackColor = &HE0E0E0
   'grd_Listad_Dsm.Row = 0
   'grd_Listad_Dsm.Col = 10
   'grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   'grd_Listad_Dsm.CellBackColor = &HE0E0E0
   grd_Listad_Dsm.Row = 0
   grd_Listad_Dsm.Col = 11
   grd_Listad_Dsm.CellAlignment = flexAlignCenterCenter
   grd_Listad_Dsm.CellBackColor = &HE0E0E0
   
   ipp_FecDsm_Dsm.Text = Format(CDate(Now), "DD/MM/YYYY")
   
   Call gs_UbiIniGrid(grd_Listad_Dsm)
End Sub
 
Private Sub txt_Comentario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End If
End Sub

Private Sub cmd_Dsm_Nuevo_Click()
   Call fs_MntBnt_Dsm(1) 'Cancelar
   cmb_FrmDsm_Dsm.ListIndex = 0
   cmd_Dsm_Insert.Tag = 1
   
   If (moddat_g_int_CodIns = CInt("000003")) Then
      'Codigo de area Tesoreria
       txt_NroDsm_Dsm.Enabled = True
       ipp_FecDsm_Dsm.Enabled = True
   Else
       txt_NroDsm_Dsm.Enabled = False
       ipp_FecDsm_Dsm.Enabled = False
   End If
    ipp_FecDsm_Dsm.Text = Format(CDate(Now), "DD/MM/YYYY")
End Sub
    
Private Sub cmd_Dsm_Borrar_Click()
   If grd_Listad_Dsm.Rows = 1 Then
      Exit Sub
   End If
   If grd_Listad_Dsm.Row = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de borrar el item ?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
       
   If (grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 12) = "I") Then
       grd_Listad_Dsm.RemoveItem (grd_Listad_Dsm.Row)
   Else
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 12) = "D"
       grd_Listad_Dsm.RowHeight(grd_Listad_Dsm.Row) = 0
   End If
   
   Call fs_sumarDesemPrmt
   Call fs_MntBnt_Dsm(4) 'Cancelar
End Sub

Private Sub cmd_Dsm_Editar_Click()
   If (grd_Listad_Dsm.Rows = 1) Then
       Exit Sub
   End If
   If (grd_Listad_Dsm.Row = 0) Then
       Exit Sub
   End If
   Call fs_MntBnt_Dsm(2) 'Editar
   
   cmd_Dsm_Insert.Tag = 2
   
   Call fs_HabFormDsm
   Call fs_mostrar_Datos
   
   'If (moddat_g_int_CodIns = CInt("000003")) Then
   '   'Codigo de area Tesoreria
   '    txt_NroDsm_Dsm.Enabled = True
   '    ipp_FecDsm_Dsm.Enabled = True
   'Else
   '    txt_NroDsm_Dsm.Enabled = False
   '    ipp_FecDsm_Dsm.Enabled = False
   'End If
End Sub

Private Sub grd_Listad_Dsm_SelChange()
   cmb_EntFin_Dsm.ListIndex = -1
   cmb_NroCta_Dsm.ListIndex = -1
   pnl_NroCCI_Dsm.Caption = ""
   cmb_FrmDsm_Dsm.ListIndex = -1
   'txt_NroDsm_Dsm.Text = ""
   'ipp_FecDsm_Dsm.Text = ""
   txt_Descrp_Dsm.Text = ""
   txt_ANombre_Dsm.Text = ""
   ipp_Import_Dsm.Text = "0.00"
   pnl_Moneda_Dsm.Caption = ""
   
   If (grd_Listad_Dsm.Rows = 1) Then
       Exit Sub
   End If
   If (grd_Listad_Dsm.Row = 0) Then
       Exit Sub
   End If
       
   Call fs_mostrar_Datos
   cmb_EntFin_Dsm.Enabled = False
   cmb_NroCta_Dsm.Enabled = False
End Sub

Private Sub fs_mostrar_Datos()
   Call gs_BuscarCombo_Item(cmb_FrmDsm_Dsm, grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 0))
   Call gs_BuscarCombo_Item(cmb_TipMto_Dsm, grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 2))
   
   
   If gs_BuscarCombo(cmb_EntFin_Dsm, Trim(grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 6))) = True Then
      cmb_EntFin_Dsm.Text = Trim(grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 6))
   Else
      cmb_EntFin_Dsm.ListIndex = -1
   End If
   
   Call cmb_EntFin_Dsm_Click
   If grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 7) <> "" Then
      If gs_BuscarCombo(cmb_NroCta_Dsm, Trim(grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 7))) = True Then
         cmb_NroCta_Dsm.Text = Trim(grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 7) & "")
      End If
   End If
   ipp_Import_Dsm.Text = grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 4)
   txt_ANombre_Dsm.Text = grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 8)
   txt_NroDsm_Dsm.Text = grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 9)
   If grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 10) = "" Then
      ipp_FecDsm_Dsm.Text = Format(CDate(Now), "DD/MM/YYYY")
   Else
      ipp_FecDsm_Dsm.Text = grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 10)
   End If
   txt_Descrp_Dsm.Text = grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 11)
   
   pnl_Moneda_Dsm.Caption = l_str_Moneda
End Sub

Function gs_BuscarCombo(p_Combo As ComboBox, p_Item As String) As Boolean
   Dim r_int_Contad  As Integer
   Dim r_int_Ubicad  As Integer
   
   r_int_Ubicad = -1
   gs_BuscarCombo = False
   
   For r_int_Contad = 0 To p_Combo.ListCount - 1
      p_Combo.ListIndex = r_int_Contad
      If Trim(p_Item) = Trim(p_Combo.Text) Then
         gs_BuscarCombo = True
         Exit For
      End If
   Next r_int_Contad
End Function

Private Sub cmd_Dsm_Cancel_Click()
   Call fs_MntBnt_Dsm(3) 'Cancelar
End Sub

Private Sub cmd_Dsm_Insert_Click()
Dim r_bol_Estado   As Boolean
Dim r_int_Fila     As Integer
Dim r_dbl_suma     As Double
Dim r_int_NumIte   As Integer

   If cmb_FrmDsm_Dsm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de desembolso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FrmDsm_Dsm)
      Exit Sub
   End If
   
   If cmb_TipMto_Dsm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de monto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMto_Dsm)
      Exit Sub
   End If
   
   
   If CDbl(Trim(ipp_Import_Dsm.Text)) = 0 Then
      MsgBox "Debe digitar un importe.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import_Dsm)
      Exit Sub
   End If
   
   If cmb_EntFin_Dsm.ListIndex = -1 Then
      MsgBox "Debe seleccionar una entidad financiera.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EntFin_Dsm)
      Exit Sub
   End If
      
   If cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 2 Then
      If cmb_NroCta_Dsm.ListIndex = -1 Then
         MsgBox "Debe seleccionar el nro de cuenta.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_NroCta_Dsm)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_ANombre_Dsm.Text)) = 0 Then
      MsgBox "Debe digitar a nombre de quien va " & IIf(UCase(Left(cmb_FrmDsm_Dsm.Text, 1)) = "T", "la ", "el ") & Trim(cmb_FrmDsm_Dsm.Text), vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ANombre_Dsm)
      Exit Sub
   End If

   If (moddat_g_int_CodIns = CInt("000003")) Then
      'Codigo de area Tesoreria
       If Len(Trim(txt_NroDsm_Dsm.Text)) = 0 Then
          'If (cmb_FrmDsm_Dsm.ListIndex = 0) Then
          '    MsgBox "Debe digitar el nro de cheque.", vbExclamation, modgen_g_str_NomPlt
          'Else
              MsgBox "Debe digitar el nro transferencia.", vbExclamation, modgen_g_str_NomPlt
          'End If
          Call gs_SetFocus(txt_NroDsm_Dsm)
          Exit Sub
       End If
      
       If Len(Trim(ipp_FecDsm_Dsm.Text)) = 0 Then
          'If (cmb_FrmDsm_Dsm.ListIndex = 0) Then
          '    MsgBox "Debe Digitar la fecha de registro del Cheque.", vbExclamation, modgen_g_str_NomPlt
          'Else
              MsgBox "Debe Digitar la fecha de registro de Transferencia.", vbExclamation, modgen_g_str_NomPlt
          'End If
          Call gs_SetFocus(ipp_FecDsm_Dsm)
          Exit Sub
       End If
   End If
      
   If (cmd_Dsm_Insert.Tag = 2) Then
       'ACTUALIZAR
       r_dbl_suma = 0
       r_dbl_suma = CDbl(grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 4))
       r_dbl_suma = CDbl(pnl_SumTot_Dsm.Caption) + CDbl(ipp_Import_Dsm.Text) - r_dbl_suma
       If (l_dbl_ImpPtm < r_dbl_suma) Then
           MsgBox "La suma de registros sobrepasa el importe del préstamo.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_Import_Dsm)
           Exit Sub
       End If
       
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 0) = cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex)
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 1) = cmb_FrmDsm_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 2) = cmb_TipMto_Dsm.ItemData(cmb_TipMto_Dsm.ListIndex)
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 3) = cmb_TipMto_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 4) = ipp_Import_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 5) = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 6) = cmb_EntFin_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 7) = cmb_NroCta_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 8) = txt_ANombre_Dsm.Text
       'grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 9) = txt_NroDsm_Dsm.Text
       'grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 10) = ipp_FecDsm_Dsm.Text
       grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 11) = txt_Descrp_Dsm.Text
       
       pnl_Moneda_Dsm.Caption = ""
           
       If (UCase(Trim(grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 12))) <> UCase(Trim("I"))) Then
           grd_Listad_Dsm.TextMatrix(grd_Listad_Dsm.Row, 12) = "U"
       End If
   Else
   'INSERTAR
       r_dbl_suma = 0
       r_dbl_suma = CDbl(pnl_SumTot_Dsm.Caption) + CDbl(ipp_Import_Dsm.Text)
       If (l_dbl_ImpPtm < r_dbl_suma) Then
           MsgBox "La suma de registros sobrepasa el importe del préstamo.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_Import_Dsm)
           Exit Sub
       End If
       
       'Genera correlativo
       r_int_NumIte = 0
       For r_int_Fila = 1 To grd_Listad_Dsm.Rows - 1
           If (r_int_NumIte <= grd_Listad_Dsm.TextMatrix(r_int_Fila, 13)) Then
               r_int_NumIte = grd_Listad_Dsm.TextMatrix(r_int_Fila, 13)
           End If
       Next
       r_int_NumIte = r_int_NumIte + 1
       
       grd_Listad_Dsm.Rows = grd_Listad_Dsm.Rows + 1
       grd_Listad_Dsm.Row = grd_Listad_Dsm.Rows - 1
       
       grd_Listad_Dsm.Col = 0
       grd_Listad_Dsm.Text = cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex)
       grd_Listad_Dsm.Col = 1
       grd_Listad_Dsm.Text = cmb_FrmDsm_Dsm.Text
       
       grd_Listad_Dsm.Col = 2
       grd_Listad_Dsm.Text = cmb_TipMto_Dsm.ItemData(cmb_TipMto_Dsm.ListIndex)
       grd_Listad_Dsm.Col = 3
       grd_Listad_Dsm.Text = cmb_TipMto_Dsm.Text
       
       grd_Listad_Dsm.Col = 4
       grd_Listad_Dsm.Text = ipp_Import_Dsm.Text
       
       grd_Listad_Dsm.Col = 5
       grd_Listad_Dsm.Text = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)
   
       grd_Listad_Dsm.Col = 6
       grd_Listad_Dsm.Text = cmb_EntFin_Dsm.Text
       
       grd_Listad_Dsm.Col = 7
       grd_Listad_Dsm.Text = cmb_NroCta_Dsm.Text
       
       grd_Listad_Dsm.Col = 8
       grd_Listad_Dsm.Text = txt_ANombre_Dsm.Text
                     
       'grd_Listad_Dsm.Col = 9
       'grd_Listad_Dsm.Text = txt_NroDsm_Dsm.Text
       'grd_Listad_Dsm.Col = 10
       'grd_Listad_Dsm.Text = ipp_FecDsm_Dsm.Text
   
       grd_Listad_Dsm.Col = 11
       grd_Listad_Dsm.Text = txt_Descrp_Dsm.Text
          
       grd_Listad_Dsm.Col = 12
       grd_Listad_Dsm.Text = "I"
       
       grd_Listad_Dsm.Col = 13
       grd_Listad_Dsm.Text = r_int_NumIte
   End If
      
   Call fs_sumarDesemPrmt
   Call fs_MntBnt_Dsm(3) 'Agregar
End Sub

Private Sub fs_sumarDesemPrmt()
Dim r_int_Fila   As Integer
Dim r_dbl_suma   As Double
    
    r_dbl_suma = 0
    For r_int_Fila = 1 To grd_Listad_Dsm.Rows - 1
        If (grd_Listad_Dsm.RowHeight(r_int_Fila) > 0) Then
            r_dbl_suma = r_dbl_suma + CDbl(grd_Listad_Dsm.TextMatrix(r_int_Fila, 4))
        End If
    Next
    pnl_SumTot_Dsm.Caption = gf_FormatoNumero(r_dbl_suma, 12, 2) & " "
    
    pnl_TotPtmo_Dsm.Caption = l_dbl_ImpPtm - r_dbl_suma
    pnl_TotPtmo_Dsm.Caption = gf_FormatoNumero(pnl_TotPtmo_Dsm.Caption, 12, 2) & " "
End Sub

Private Sub fs_MntBnt_Dsm(p_Tipo As Integer)
'Desabilitar = 0; Nuevo = 1; Editar = 2; Agregar = 3; Cancelar = 4
   If (p_Tipo = 0) Then '---desabilitar----
       cmd_Dsm_Nuevo.Enabled = False
       cmd_Dsm_Borrar.Enabled = False
       cmd_Dsm_Editar.Enabled = False
       cmd_Dsm_Insert.Enabled = False
       cmd_Dsm_Cancel.Enabled = False
       cmb_NroCta_Dsm.ListIndex = -1
       pnl_NroCCI_Dsm.Caption = ""
       cmb_FrmDsm_Dsm.ListIndex = -1
       cmb_TipMto_Dsm.ListIndex = -1
       cmb_EntFin_Dsm.ListIndex = -1
       'txt_NroDsm_Dsm.Text = ""
       'ipp_FecDsm_Dsm.Text = ""
       txt_Descrp_Dsm.Text = ""
       txt_ANombre_Dsm.Text = ""
       ipp_Import_Dsm.Text = "0.00"
       pnl_Moneda_Dsm.Caption = l_str_Moneda
       cmb_EntFin_Dsm.Enabled = False
       cmb_NroCta_Dsm.Enabled = False
       cmb_FrmDsm_Dsm.Enabled = False
       cmb_TipMto_Dsm.Enabled = False
       'txt_NroDsm_Dsm.Enabled = False
       'ipp_FecDsm_Dsm.Enabled = False
       txt_Descrp_Dsm.Enabled = False
       txt_ANombre_Dsm.Enabled = False
       ipp_Import_Dsm.Enabled = False
   ElseIf (p_Tipo = 1) Then '---nuevo----
       cmd_Dsm_Nuevo.Enabled = False
       cmd_Dsm_Borrar.Enabled = False
       cmd_Dsm_Editar.Enabled = False
       cmd_Dsm_Insert.Enabled = True
       cmd_Dsm_Cancel.Enabled = True
       cmb_NroCta_Dsm.ListIndex = -1
       pnl_NroCCI_Dsm.Caption = ""
       cmb_FrmDsm_Dsm.ListIndex = -1
       cmb_TipMto_Dsm.ListIndex = -1
       cmb_EntFin_Dsm.ListIndex = -1
       'txt_NroDsm_Dsm.Text = ""
       'ipp_FecDsm_Dsm.Text = ""
       txt_Descrp_Dsm.Text = ""
       txt_ANombre_Dsm.Text = ""
       ipp_Import_Dsm.Text = "0.00"
       pnl_Moneda_Dsm.Caption = ""
       cmb_EntFin_Dsm.Enabled = True
       cmb_NroCta_Dsm.Enabled = True
       cmb_FrmDsm_Dsm.Enabled = True
       cmb_TipMto_Dsm.Enabled = True
       'txt_NroDsm_Dsm.Enabled = True
       'ipp_FecDsm_Dsm.Enabled = True
       txt_Descrp_Dsm.Enabled = True
       txt_ANombre_Dsm.Enabled = True
       ipp_Import_Dsm.Enabled = True
       pnl_Moneda_Dsm.Caption = l_str_Moneda
       grd_Listad_Dsm.Enabled = False
       Call gs_UbiIniGrid(grd_Listad_Dsm)
       Call gs_SetFocus(cmb_FrmDsm_Dsm)
   ElseIf (p_Tipo = 2) Then '---Editar-----
       cmd_Dsm_Nuevo.Enabled = False
       cmd_Dsm_Borrar.Enabled = False
       cmd_Dsm_Editar.Enabled = False
       cmd_Dsm_Insert.Enabled = True
       cmd_Dsm_Cancel.Enabled = True
       cmb_EntFin_Dsm.Enabled = True
       cmb_NroCta_Dsm.Enabled = True
       cmb_FrmDsm_Dsm.Enabled = True
       cmb_TipMto_Dsm.Enabled = True
       'txt_NroDsm_Dsm.Enabled = True
       'ipp_FecDsm_Dsm.Enabled = True
       txt_Descrp_Dsm.Enabled = True
       txt_ANombre_Dsm.Enabled = True
       ipp_Import_Dsm.Enabled = True
       pnl_Moneda_Dsm.Caption = l_str_Moneda
       grd_Listad_Dsm.Enabled = False
       Call gs_SetFocus(cmb_FrmDsm_Dsm)
   ElseIf (p_Tipo = 3) Then '---Agregar-----
       cmd_Dsm_Nuevo.Enabled = True
       cmd_Dsm_Borrar.Enabled = True
       cmd_Dsm_Editar.Enabled = True
       cmd_Dsm_Insert.Enabled = False
       cmd_Dsm_Cancel.Enabled = False
       cmb_EntFin_Dsm.ListIndex = -1
       cmb_NroCta_Dsm.ListIndex = -1
       pnl_NroCCI_Dsm.Caption = ""
       cmb_FrmDsm_Dsm.ListIndex = -1
       cmb_TipMto_Dsm.ListIndex = -1
       'txt_NroDsm_Dsm.Text = ""
       'ipp_FecDsm_Dsm.Text = ""
       txt_Descrp_Dsm.Text = ""
       txt_ANombre_Dsm.Text = ""
       ipp_Import_Dsm.Text = "0.00"
       pnl_Moneda_Dsm.Caption = ""
       cmb_EntFin_Dsm.Enabled = False
       cmb_NroCta_Dsm.Enabled = False
       cmb_FrmDsm_Dsm.Enabled = False
       cmb_TipMto_Dsm.Enabled = False
       'txt_NroDsm_Dsm.Enabled = False
       'ipp_FecDsm_Dsm.Enabled = False
       txt_Descrp_Dsm.Enabled = False
       txt_ANombre_Dsm.Enabled = False
       ipp_Import_Dsm.Enabled = False
       grd_Listad_Dsm.Enabled = True
       Call gs_UbiIniGrid(grd_Listad_Dsm)
       Call gs_SetFocus(cmd_Dsm_Nuevo)
   ElseIf (p_Tipo = 4) Then '---Cancelar-----
       cmd_Dsm_Nuevo.Enabled = True
       cmd_Dsm_Borrar.Enabled = True
       cmd_Dsm_Editar.Enabled = True
       cmd_Dsm_Insert.Enabled = False
       cmd_Dsm_Cancel.Enabled = False
       cmb_EntFin_Dsm.ListIndex = -1
       cmb_NroCta_Dsm.ListIndex = -1
       pnl_NroCCI_Dsm.Caption = ""
       cmb_FrmDsm_Dsm.ListIndex = -1
       cmb_TipMto_Dsm.ListIndex = -1
       'txt_NroDsm_Dsm.Text = ""
       'ipp_FecDsm_Dsm.Text = ""
       txt_Descrp_Dsm.Text = ""
       txt_ANombre_Dsm.Text = ""
       ipp_Import_Dsm.Text = "0.00"
       pnl_Moneda_Dsm.Caption = ""
       cmb_EntFin_Dsm.Enabled = False
       cmb_NroCta_Dsm.Enabled = False
       cmb_FrmDsm_Dsm.Enabled = False
       cmb_TipMto_Dsm.Enabled = False
       'txt_NroDsm_Dsm.Enabled = False
       'ipp_FecDsm_Dsm.Enabled = False
       txt_Descrp_Dsm.Enabled = False
       txt_ANombre_Dsm.Enabled = False
       ipp_Import_Dsm.Enabled = False
       grd_Listad_Dsm.Enabled = True
       Call gs_UbiIniGrid(grd_Listad_Dsm)
   End If
End Sub

Private Sub cmb_FrmDsm_Dsm_Click()
    'If (cmb_FrmDsm_Dsm.ListIndex <> -1) Then
    '    lbl_NumDsm_Dsm.Caption = "Nro " & UCase(Left(Trim(cmb_FrmDsm_Dsm.Text), 1)) & LCase(Right(Trim(cmb_FrmDsm_Dsm.Text), Len(Trim(cmb_FrmDsm_Dsm.Text)) - 1))
    '    lbl_FchDsm_Dsm.Caption = "Fecha Reg " & UCase(Left(Trim(cmb_FrmDsm_Dsm.Text), 1)) & LCase(Right(Left(Trim(cmb_FrmDsm_Dsm.Text), 4), 3))
        
    '    lbl_NumDsm_Dsm.Caption = lbl_NumDsm_Dsm.Caption & ":"
    '    lbl_FchDsm_Dsm.Caption = lbl_FchDsm_Dsm.Caption & ":"
    'Else
    '    lbl_NumDsm_Dsm.Caption = "Nro Desembolso:"
    '    lbl_FchDsm_Dsm.Caption = "Fecha Reg Dsm:"
    'End If
    Call fs_HabFormDsm
End Sub

Private Sub fs_HabFormDsm()
    If cmd_Dsm_Editar.Enabled = False And cmd_Grabar.Enabled = True Then
       If cmb_FrmDsm_Dsm.ListIndex > -1 Then
          If cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 1 Or cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 3 Or _
             cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 4 Then
             cmb_EntFin_Dsm.ListIndex = -1
             cmb_NroCta_Dsm.ListIndex = -1
             pnl_NroCCI_Dsm.Caption = ""
             cmb_NroCta_Dsm.Enabled = False
          Else
             cmb_EntFin_Dsm.ListIndex = -1
             cmb_EntFin_Dsm.Enabled = True
             cmb_NroCta_Dsm.Enabled = True
          End If
       End If
    End If
End Sub

Private Sub cmd_Dsm_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmb_EntFin_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_NroCta_Dsm.Enabled = False Then
          Call gs_SetFocus(txt_ANombre_Dsm)
      Else
          Call gs_SetFocus(cmb_NroCta_Dsm)
      End If
   End If
End Sub

Private Sub cmb_NroCta_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ANombre_Dsm)
   End If
End Sub

Private Sub cmb_FrmDsm_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipMto_Dsm)
   End If
End Sub

Private Sub ipp_Import_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_EntFin_Dsm.Enabled = False Then
         Call gs_SetFocus(txt_ANombre_Dsm)
      Else
         Call gs_SetFocus(cmb_EntFin_Dsm)
      End If
   End If
End Sub

Private Sub txt_ANombre_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Descrp_Dsm)
   Else
      'KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & " '")
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- _.")
   End If
End Sub

Private Sub txt_Descrp_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_NroDsm_Dsm.Visible = False) Then
          Call gs_SetFocus(cmd_Dsm_Insert)
      Else
          Call gs_SetFocus(txt_NroDsm_Dsm)
      End If
   Else
      'KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & " '")
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub txt_NroDsm_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecDsm_Dsm)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-'")
   End If
End Sub

Private Sub ipp_FecDsm_Dsm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Comentario)
   End If
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel   As Excel.Application
Dim r_int_Filaux  As Integer
Dim r_int_filExl  As Integer
Dim r_int_totExl  As Integer
Dim r_str_Cadena  As String
                
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
        'Unir celdas
        .Range("B2") = "NRO OPERACION:"
        .Range("B3") = "CLIENTE:"
        .Range("B4") = "PRODUCTO:"
        .Range("B5") = "PROYECTO:"
        .Range("B6") = "PROMOTOR:"
        
        .Range("C2") = pnl_NumOpe.Caption
        .Range("C3") = Trim(pnl_NomCli.Caption)
        .Range("C4") = Trim(pnl_Produc.Caption)
        .Range("C5") = Trim(pnl_Prycto_Dsm.Caption)
        .Range("C6") = Trim(l_str_Prmtor)
        
        If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
           'Legal 1 y Operaciones 1
           r_int_totExl = 8
           r_str_Cadena = "H"
        Else
           r_int_totExl = 10
           r_str_Cadena = "J"
        End If
        
        r_int_filExl = 8
        .Range("B" & r_int_filExl) = "DATOS DE DESEMBOLSO A PROMOTOR"
        .Range("B" & r_int_filExl & ":" & r_str_Cadena & r_int_filExl).Font.Bold = True
        .Range("B" & r_int_filExl & ":" & r_str_Cadena & r_int_filExl).Merge
        .Range("B" & r_int_filExl).HorizontalAlignment = xlHAlignCenter
        
        r_int_filExl = r_int_filExl + 1
        .Columns("G").HorizontalAlignment = xlHAlignLeft
        .Range("B" & r_int_filExl & ":" & r_str_Cadena & r_int_filExl).HorizontalAlignment = xlHAlignCenter
        .Range("B" & r_int_filExl & ":" & r_str_Cadena & r_int_filExl).Font.Bold = True
        .Range("B" & r_int_filExl & ":" & r_str_Cadena & r_int_filExl).Interior.Color = RGB(146, 208, 80)
        
        For r_int_Filaux = 2 To r_int_totExl
            .Cells(r_int_filExl, r_int_Filaux).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Cells(r_int_filExl, r_int_Filaux).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Cells(r_int_filExl, r_int_Filaux).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Cells(r_int_filExl, r_int_Filaux).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Next
                    
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 14
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 26
        .Columns("E").ColumnWidth = 18
        .Columns("F").ColumnWidth = 13
        .Columns("G").ColumnWidth = 30
        If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
           'Legal 1 y Operaciones 1
           .Columns("H").ColumnWidth = 30
        Else
           .Columns("H").ColumnWidth = 30
           .Columns("I").ColumnWidth = 11
           .Columns("J").ColumnWidth = 42
        End If
              
        .Cells(r_int_filExl, 2) = "Forma Pago"
        .Cells(r_int_filExl, 3) = "Tipo Monto"
        .Cells(r_int_filExl, 4) = "Banco"
        .Cells(r_int_filExl, 5) = "Nro Cuenta"
        .Cells(r_int_filExl, 6) = "Importe"
        .Cells(r_int_filExl, 7) = "A Nombre de"
        If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
           'Legal 1 y Operaciones 1
           .Cells(r_int_filExl, 8) = "Descripción"
        Else
           .Cells(r_int_filExl, 8) = "Nro Desembolso"
           .Cells(r_int_filExl, 9) = "Fecha Reg."
           .Cells(r_int_filExl, 10) = "Descripción"
        End If
                
         For r_int_Filaux = 1 To grd_Listad_Dsm.Rows - 1
             r_int_filExl = r_int_filExl + 1
             .Cells(r_int_filExl, 2).NumberFormat = "@"
             .Cells(r_int_filExl, 3).NumberFormat = "@"
             .Cells(r_int_filExl, 5).NumberFormat = "@"
             .Cells(r_int_filExl, 6).NumberFormat = "###,###,##0.00" '"@"
             .Cells(r_int_filExl, 7).NumberFormat = "@"
             .Cells(r_int_filExl, 8).NumberFormat = "@"
             .Cells(r_int_filExl, 9).NumberFormat = "@"
             .Cells(r_int_filExl, 10).NumberFormat = "@"
             
             .Cells(r_int_filExl, 2) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 1)
             .Cells(r_int_filExl, 3) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 3)
             .Cells(r_int_filExl, 4) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 6)
             .Cells(r_int_filExl, 5) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 7)
             .Cells(r_int_filExl, 6) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 4)
             .Cells(r_int_filExl, 7) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 8)
             If (moddat_g_int_CodIns = CInt("000001") Or moddat_g_int_CodIns = CInt("000002") Or moddat_g_int_CodIns = 0) Then
                'Legal 1 y Operaciones 1
                .Cells(r_int_filExl, 8) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 11)
             Else
                .Cells(r_int_filExl, 8) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 9)
                .Cells(r_int_filExl, 9) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 10)
                .Cells(r_int_filExl, 10) = grd_Listad_Dsm.TextMatrix(r_int_Filaux, 11)
             End If
         Next
         
         r_int_filExl = r_int_filExl + 1
         .Cells(r_int_filExl, 5) = "Suma Total ==>"
         .Cells(r_int_filExl, 6) = pnl_SumTot_Dsm.Caption
         .Range("E" & r_int_filExl & ":F" & r_int_filExl).Interior.Color = RGB(146, 208, 80)
         .Range("E" & r_int_filExl & ":F" & r_int_filExl).Font.Bold = True
         
         .Range("A1:J" & r_int_filExl).Font.Name = "Arial"
         .Range("A1:J" & r_int_filExl).Font.Size = 8
   End With

   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_EntFin_Dsm_Click()
Dim r_int_Fila As Integer

    cmb_NroCta_Dsm.Clear
    txt_ANombre_Dsm.Text = ""
    If cmb_FrmDsm_Dsm.ListIndex > -1 Then
       If cmb_FrmDsm_Dsm.ItemData(cmb_FrmDsm_Dsm.ListIndex) = 2 Then
          'transferencia
          For r_int_Fila = 1 To UBound(l_arr_CtaBco)
              If (cmd_Dsm_Insert.Enabled = False) Then
                  If (l_arr_CtaBco(r_int_Fila).Genera_Codigo = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)) Then
                      cmb_NroCta_Dsm.AddItem (Trim(l_arr_CtaBco(r_int_Fila).Genera_Nombre))
                  End If
              Else
                  If (cmd_Dsm_Insert.Tag = 1) Then
                      If (l_arr_CtaBco(r_int_Fila).Genera_FlgAso = 1 And _
                          l_arr_CtaBco(r_int_Fila).Genera_Codigo = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)) Then
                          cmb_NroCta_Dsm.AddItem (Trim(l_arr_CtaBco(r_int_Fila).Genera_Nombre))
                      End If
                  Else
                      If (l_arr_CtaBco(r_int_Fila).Genera_Codigo = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)) Then
                          cmb_NroCta_Dsm.AddItem (Trim(l_arr_CtaBco(r_int_Fila).Genera_Nombre))
                      End If
                  End If
              End If
          Next
       Else
          'cheque
          For r_int_Fila = 1 To UBound(l_arr_CtaBco)
              If (l_arr_CtaBco(r_int_Fila).Genera_Codigo = CStr(l_arr_Bancos(cmb_EntFin_Dsm.ListIndex + 1).Genera_Codigo)) Then
                  txt_ANombre_Dsm.Text = Trim(l_arr_CtaBco(r_int_Fila).Genera_ConHip)
                  Exit For
              End If
          Next
       End If
    End If
End Sub

