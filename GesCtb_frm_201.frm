VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_GesPer_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16125
   Icon            =   "GesCtb_frm_201.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   16125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9030
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16140
      _Version        =   65536
      _ExtentX        =   28469
      _ExtentY        =   15928
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
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   16020
         _Version        =   65536
         _ExtentX        =   28257
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
            Left            =   660
            TabIndex        =   2
            Top             =   30
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Gestión al Personal"
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
            Picture         =   "GesCtb_frm_201.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   8205
         Left            =   60
         TabIndex        =   3
         Top             =   780
         Width           =   16020
         _Version        =   65536
         _ExtentX        =   28257
         _ExtentY        =   14473
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
         Begin TabDlg.SSTab tab_GesPer 
            Height          =   8145
            Left            =   30
            TabIndex        =   4
            Top             =   30
            Width           =   15975
            _ExtentX        =   28178
            _ExtentY        =   14367
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Gestion de Pagos"
            TabPicture(0)   =   "GesCtb_frm_201.frx":0316
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel3"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel5"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel9"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel2"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Gestion de Vacaciones"
            TabPicture(1)   =   "GesCtb_frm_201.frx":0332
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel33"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "SSPanel21"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "SSPanel15"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).ControlCount=   3
            TabCaption(2)   =   "Gestion de Autorizaciones"
            TabPicture(2)   =   "GesCtb_frm_201.frx":034E
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SSPanel19"
            Tab(2).Control(1)=   "SSPanel20"
            Tab(2).Control(2)=   "SSPanel42"
            Tab(2).ControlCount=   3
            Begin Threed.SSPanel SSPanel2 
               Height          =   645
               Left            =   30
               TabIndex        =   5
               Top             =   360
               Width           =   15885
               _Version        =   65536
               _ExtentX        =   28028
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
               Begin VB.CommandButton cmd_GenPag 
                  Appearance      =   0  'Flat
                  Height          =   585
                  Left            =   3660
                  Picture         =   "GesCtb_frm_201.frx":036A
                  Style           =   1  'Graphical
                  TabIndex        =   13
                  ToolTipText     =   "Generar Registros"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_AgregaPag 
                  Height          =   585
                  Left            =   1230
                  Picture         =   "GesCtb_frm_201.frx":0674
                  Style           =   1  'Graphical
                  TabIndex        =   12
                  ToolTipText     =   "Adicionar"
                  Top             =   30
                  Width           =   615
               End
               Begin VB.CommandButton cmd_BorrarPag 
                  Height          =   585
                  Left            =   1860
                  Picture         =   "GesCtb_frm_201.frx":097E
                  Style           =   1  'Graphical
                  TabIndex        =   11
                  ToolTipText     =   "Eliminar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_SalirPag 
                  Height          =   585
                  Left            =   15270
                  Picture         =   "GesCtb_frm_201.frx":0C88
                  Style           =   1  'Graphical
                  TabIndex        =   10
                  ToolTipText     =   "Salir"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_BusPag 
                  Height          =   585
                  Left            =   30
                  Picture         =   "GesCtb_frm_201.frx":10CA
                  Style           =   1  'Graphical
                  TabIndex        =   9
                  ToolTipText     =   "Buscar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_LimPag 
                  Height          =   585
                  Left            =   630
                  Picture         =   "GesCtb_frm_201.frx":13D4
                  Style           =   1  'Graphical
                  TabIndex        =   8
                  ToolTipText     =   "Limpiar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_ConsulPag 
                  Height          =   585
                  Left            =   2460
                  Picture         =   "GesCtb_frm_201.frx":16DE
                  Style           =   1  'Graphical
                  TabIndex        =   7
                  ToolTipText     =   "Consultar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_ExpPag 
                  Height          =   585
                  Left            =   3060
                  Picture         =   "GesCtb_frm_201.frx":19E8
                  Style           =   1  'Graphical
                  TabIndex        =   6
                  ToolTipText     =   "Exportar a Excel"
                  Top             =   30
                  Width           =   585
               End
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   825
               Left            =   30
               TabIndex        =   14
               Top             =   1050
               Width           =   15885
               _Version        =   65536
               _ExtentX        =   28028
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
               Begin VB.ComboBox cmb_SucPag 
                  Height          =   315
                  Left            =   1170
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   420
                  Width           =   3465
               End
               Begin VB.ComboBox cmb_EmpPag 
                  Height          =   315
                  Left            =   1170
                  Style           =   2  'Dropdown List
                  TabIndex        =   15
                  Top             =   90
                  Width           =   3465
               End
               Begin EditLib.fpDateTime ipp_FecIniPag 
                  Height          =   315
                  Left            =   6780
                  TabIndex        =   17
                  Top             =   420
                  Width           =   1365
                  _Version        =   196608
                  _ExtentX        =   2408
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
                  Text            =   "28/09/2004"
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
               Begin EditLib.fpDateTime ipp_FecFinPag 
                  Height          =   315
                  Left            =   8160
                  TabIndex        =   18
                  Top             =   420
                  Width           =   1365
                  _Version        =   196608
                  _ExtentX        =   2408
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
                  Text            =   "28/09/2004"
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
               Begin Threed.SSPanel pnl_PerPag 
                  Height          =   315
                  Left            =   6780
                  TabIndex        =   19
                  Top             =   90
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
                  Alignment       =   1
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha:"
                  Height          =   195
                  Left            =   5310
                  TabIndex        =   23
                  Top             =   450
                  Width           =   495
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Sucursal:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   22
                  Top             =   450
                  Width           =   660
               End
               Begin VB.Label lbl_NomEti 
                  AutoSize        =   -1  'True
                  Caption         =   "Empresa:"
                  Height          =   195
                  Index           =   0
                  Left            =   180
                  TabIndex        =   21
                  Top             =   120
                  Width           =   660
               End
               Begin VB.Label lbl_NomEti 
                  AutoSize        =   -1  'True
                  Caption         =   "Período Vigente:"
                  Height          =   195
                  Index           =   2
                  Left            =   5310
                  TabIndex        =   20
                  Top             =   120
                  Width           =   1200
               End
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   5490
               Left            =   30
               TabIndex        =   24
               Top             =   1920
               Width           =   15885
               _Version        =   65536
               _ExtentX        =   28019
               _ExtentY        =   9684
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
               Begin Threed.SSPanel SSPanel4 
                  Height          =   285
                  Left            =   8115
                  TabIndex        =   25
                  Top             =   60
                  Width           =   1080
                  _Version        =   65536
                  _ExtentX        =   1905
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
               Begin Threed.SSPanel SSPanel17 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   26
                  Top             =   60
                  Width           =   1275
                  _Version        =   65536
                  _ExtentX        =   2240
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
               Begin MSFlexGridLib.MSFlexGrid grd_ListPag 
                  Height          =   5100
                  Left            =   30
                  TabIndex        =   27
                  Top             =   360
                  Width           =   15840
                  _ExtentX        =   27940
                  _ExtentY        =   8996
                  _Version        =   393216
                  Rows            =   30
                  Cols            =   20
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel pnl_DebMN 
                  Height          =   285
                  Left            =   9180
                  TabIndex        =   28
                  Top             =   60
                  Width           =   885
                  _Version        =   65536
                  _ExtentX        =   1570
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
               Begin Threed.SSPanel pnl_HabME 
                  Height          =   285
                  Left            =   10050
                  TabIndex        =   29
                  Top             =   60
                  Width           =   1155
                  _Version        =   65536
                  _ExtentX        =   2046
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
               Begin Threed.SSPanel SSPanel13 
                  Height          =   285
                  Left            =   11190
                  TabIndex        =   30
                  Top             =   60
                  Width           =   1065
                  _Version        =   65536
                  _ExtentX        =   1879
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Contabilizado"
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
                  Left            =   14340
                  TabIndex        =   31
                  Top             =   60
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2081
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   " Seleccionar"
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
                  Begin VB.CheckBox chkSelectPag 
                     BackColor       =   &H00004000&
                     Caption         =   "Check1"
                     Height          =   255
                     Left            =   940
                     TabIndex        =   32
                     Top             =   0
                     Width           =   255
                  End
               End
               Begin Threed.SSPanel SSPanel16 
                  Height          =   285
                  Left            =   2460
                  TabIndex        =   33
                  Top             =   60
                  Width           =   2415
                  _Version        =   65536
                  _ExtentX        =   4251
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Tipo Operación"
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
               Begin Threed.SSPanel SSPanel18 
                  Height          =   285
                  Left            =   60
                  TabIndex        =   34
                  Top             =   60
                  Width           =   1160
                  _Version        =   65536
                  _ExtentX        =   2046
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
               Begin Threed.SSPanel SSPanel10 
                  Height          =   285
                  Left            =   12240
                  TabIndex        =   35
                  Top             =   60
                  Width           =   1080
                  _Version        =   65536
                  _ExtentX        =   1905
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
               Begin Threed.SSPanel SSPanel11 
                  Height          =   285
                  Left            =   13290
                  TabIndex        =   36
                  Top             =   60
                  Width           =   1080
                  _Version        =   65536
                  _ExtentX        =   1905
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
               Begin Threed.SSPanel SSPanel8 
                  Height          =   285
                  Left            =   4860
                  TabIndex        =   37
                  Top             =   60
                  Width           =   3270
                  _Version        =   65536
                  _ExtentX        =   5768
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Trabajador"
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   670
               Left            =   -74970
               TabIndex        =   38
               Top             =   360
               Width           =   15885
               _Version        =   65536
               _ExtentX        =   28019
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
               Begin VB.CommandButton cmd_EditVac 
                  Height          =   585
                  Left            =   1230
                  Picture         =   "GesCtb_frm_201.frx":1CF2
                  Style           =   1  'Graphical
                  TabIndex        =   60
                  ToolTipText     =   "Modificar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_DetVac 
                  Height          =   585
                  Left            =   1830
                  Picture         =   "GesCtb_frm_201.frx":1FFC
                  Style           =   1  'Graphical
                  TabIndex        =   61
                  ToolTipText     =   "Detalle"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_ExpVac 
                  Height          =   585
                  Left            =   2430
                  Picture         =   "GesCtb_frm_201.frx":243E
                  Style           =   1  'Graphical
                  TabIndex        =   62
                  ToolTipText     =   "Exportar a Excel"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_LimVac 
                  Height          =   585
                  Left            =   630
                  Picture         =   "GesCtb_frm_201.frx":2748
                  Style           =   1  'Graphical
                  TabIndex        =   59
                  ToolTipText     =   "Limpiar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_BusVac 
                  Height          =   585
                  Left            =   30
                  Picture         =   "GesCtb_frm_201.frx":2A52
                  Style           =   1  'Graphical
                  TabIndex        =   58
                  ToolTipText     =   "Buscar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_SalirVac 
                  Height          =   585
                  Left            =   15270
                  Picture         =   "GesCtb_frm_201.frx":2D5C
                  Style           =   1  'Graphical
                  TabIndex        =   39
                  ToolTipText     =   "Salir"
                  Top             =   30
                  Width           =   585
               End
            End
            Begin Threed.SSPanel SSPanel21 
               Height          =   6330
               Left            =   -74970
               TabIndex        =   40
               Top             =   1080
               Width           =   15885
               _Version        =   65536
               _ExtentX        =   28019
               _ExtentY        =   11165
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
               Begin Threed.SSPanel pln_Vencido_Aut 
                  Height          =   285
                  Left            =   10690
                  TabIndex        =   46
                  Top             =   60
                  Width           =   1500
                  _Version        =   65536
                  _ExtentX        =   2646
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Vencidos (Días)"
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
               Begin Threed.SSPanel pln_FecIng_Aut 
                  Height          =   285
                  Left            =   9135
                  TabIndex        =   41
                  Top             =   60
                  Width           =   1570
                  _Version        =   65536
                  _ExtentX        =   2769
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Fecha Ingreso"
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
               Begin Threed.SSPanel pln_NroDoc_Aut 
                  Height          =   285
                  Left            =   1290
                  TabIndex        =   42
                  Top             =   60
                  Width           =   1755
                  _Version        =   65536
                  _ExtentX        =   3096
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
               Begin MSFlexGridLib.MSFlexGrid grd_ListVac 
                  Height          =   5940
                  Left            =   40
                  TabIndex        =   43
                  Top             =   360
                  Width           =   15840
                  _ExtentX        =   27940
                  _ExtentY        =   10478
                  _Version        =   393216
                  Rows            =   30
                  Cols            =   14
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel pln_DiaGoz_Aut 
                  Height          =   285
                  Left            =   12180
                  TabIndex        =   44
                  Top             =   60
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Gozados (Días)"
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
               Begin Threed.SSPanel pnl_CodPla_Aut 
                  Height          =   285
                  Left            =   60
                  TabIndex        =   45
                  Top             =   60
                  Width           =   1260
                  _Version        =   65536
                  _ExtentX        =   2222
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Código Planilla"
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
               Begin Threed.SSPanel pln_SldVen_Aut 
                  Height          =   285
                  Left            =   13740
                  TabIndex        =   47
                  Top             =   60
                  Width           =   1800
                  _Version        =   65536
                  _ExtentX        =   3175
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Saldo (Días)"
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
               Begin Threed.SSPanel pln_NomTra_Aut 
                  Height          =   285
                  Left            =   3030
                  TabIndex        =   48
                  Top             =   60
                  Width           =   4720
                  _Version        =   65536
                  _ExtentX        =   8326
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Trabajador"
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
               Begin Threed.SSPanel pln_Situac_Aut 
                  Height          =   285
                  Left            =   7740
                  TabIndex        =   94
                  Top             =   60
                  Width           =   1425
                  _Version        =   65536
                  _ExtentX        =   2514
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Situación"
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
            Begin Threed.SSPanel SSPanel3 
               Height          =   630
               Left            =   30
               TabIndex        =   49
               Top             =   7460
               Width           =   15885
               _Version        =   65536
               _ExtentX        =   28019
               _ExtentY        =   1111
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
               Begin VB.ComboBox cmb_BusPag 
                  Height          =   315
                  Left            =   1620
                  Style           =   2  'Dropdown List
                  TabIndex        =   51
                  Top             =   180
                  Width           =   2595
               End
               Begin VB.TextBox txt_BusPag 
                  Height          =   315
                  Left            =   5400
                  MaxLength       =   100
                  TabIndex        =   50
                  Top             =   180
                  Width           =   4425
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Buscar Por:"
                  Height          =   195
                  Left            =   4530
                  TabIndex        =   53
                  Top             =   240
                  Width           =   825
               End
               Begin VB.Label lbl_NomEti 
                  AutoSize        =   -1  'True
                  Caption         =   "Columna a Buscar:"
                  Height          =   195
                  Index           =   1
                  Left            =   180
                  TabIndex        =   52
                  Top             =   240
                  Width           =   1335
               End
            End
            Begin Threed.SSPanel SSPanel33 
               Height          =   630
               Left            =   -74970
               TabIndex        =   54
               Top             =   7460
               Width           =   15885
               _Version        =   65536
               _ExtentX        =   28019
               _ExtentY        =   1111
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
               Begin VB.TextBox txt_BusVac 
                  Height          =   315
                  Left            =   5400
                  MaxLength       =   100
                  TabIndex        =   56
                  Top             =   180
                  Width           =   4425
               End
               Begin VB.ComboBox cmb_BusVac 
                  Height          =   315
                  Left            =   1620
                  Style           =   2  'Dropdown List
                  TabIndex        =   55
                  Top             =   180
                  Width           =   2595
               End
               Begin VB.Label lbl_NomEti 
                  AutoSize        =   -1  'True
                  Caption         =   "Columna a Buscar:"
                  Height          =   195
                  Index           =   5
                  Left            =   180
                  TabIndex        =   63
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Buscar Por:"
                  Height          =   195
                  Left            =   4500
                  TabIndex        =   57
                  Top             =   240
                  Width           =   825
               End
            End
            Begin Threed.SSPanel SSPanel19 
               Height          =   670
               Left            =   -74970
               TabIndex        =   64
               Top             =   360
               Width           =   15885
               _Version        =   65536
               _ExtentX        =   28019
               _ExtentY        =   1182
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
               Begin VB.CommandButton cmd_BusAut 
                  Height          =   585
                  Left            =   30
                  Picture         =   "GesCtb_frm_201.frx":319E
                  Style           =   1  'Graphical
                  TabIndex        =   92
                  ToolTipText     =   "Buscar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_LimAut 
                  Height          =   585
                  Left            =   630
                  Picture         =   "GesCtb_frm_201.frx":34A8
                  Style           =   1  'Graphical
                  TabIndex        =   91
                  ToolTipText     =   "Limpiar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_Reversa 
                  Height          =   585
                  Left            =   2430
                  Picture         =   "GesCtb_frm_201.frx":37B2
                  Style           =   1  'Graphical
                  TabIndex        =   82
                  ToolTipText     =   "Reversa"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_AprAut 
                  Height          =   585
                  Left            =   1230
                  Picture         =   "GesCtb_frm_201.frx":3ABC
                  Style           =   1  'Graphical
                  TabIndex        =   65
                  ToolTipText     =   "Aprobar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_RhzAut 
                  Height          =   585
                  Left            =   1830
                  Picture         =   "GesCtb_frm_201.frx":3DC6
                  Style           =   1  'Graphical
                  TabIndex        =   66
                  ToolTipText     =   "Rechazar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_SalirAut 
                  Height          =   600
                  Left            =   15270
                  Picture         =   "GesCtb_frm_201.frx":4208
                  Style           =   1  'Graphical
                  TabIndex        =   67
                  ToolTipText     =   "Salir de la Opción"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_ConsulAut 
                  Height          =   585
                  Left            =   3030
                  Picture         =   "GesCtb_frm_201.frx":464A
                  Style           =   1  'Graphical
                  TabIndex        =   68
                  ToolTipText     =   "Consultar"
                  Top             =   30
                  Width           =   585
               End
               Begin VB.CommandButton cmd_ExpAut 
                  Height          =   585
                  Left            =   3630
                  Picture         =   "GesCtb_frm_201.frx":4954
                  Style           =   1  'Graphical
                  TabIndex        =   70
                  ToolTipText     =   "Exportar a Excel"
                  Top             =   30
                  Width           =   585
               End
            End
            Begin Threed.SSPanel SSPanel20 
               Height          =   6300
               Left            =   -74970
               TabIndex        =   69
               Top             =   1770
               Width           =   15885
               _Version        =   65536
               _ExtentX        =   28019
               _ExtentY        =   11112
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
               Begin Threed.SSPanel SSPanel35 
                  Height          =   285
                  Left            =   8535
                  TabIndex        =   75
                  Top             =   60
                  Width           =   1050
                  _Version        =   65536
                  _ExtentX        =   1852
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Fecha Hasta"
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
               Begin Threed.SSPanel SSPanel27 
                  Height          =   285
                  Left            =   1245
                  TabIndex        =   71
                  Top             =   60
                  Width           =   1330
                  _Version        =   65536
                  _ExtentX        =   2346
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Fecha Operación"
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
               Begin Threed.SSPanel SSPanel34 
                  Height          =   285
                  Left            =   7485
                  TabIndex        =   74
                  Top             =   60
                  Width           =   1065
                  _Version        =   65536
                  _ExtentX        =   1870
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Fecha Desde"
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
               Begin MSFlexGridLib.MSFlexGrid grd_ListAut 
                  Height          =   5900
                  Left            =   30
                  TabIndex        =   72
                  Top             =   360
                  Width           =   15840
                  _ExtentX        =   27940
                  _ExtentY        =   10398
                  _Version        =   393216
                  Rows            =   30
                  Cols            =   14
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel36 
                  Height          =   285
                  Left            =   60
                  TabIndex        =   76
                  Top             =   60
                  Width           =   1200
                  _Version        =   65536
                  _ExtentX        =   2117
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Código Interno"
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
               Begin Threed.SSPanel SSPanel37 
                  Height          =   285
                  Left            =   9570
                  TabIndex        =   77
                  Top             =   60
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1288
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Días Sol."
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
               Begin Threed.SSPanel SSPanel38 
                  Height          =   285
                  Left            =   10290
                  TabIndex        =   78
                  Top             =   60
                  Width           =   3165
                  _Version        =   65536
                  _ExtentX        =   5583
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Comentario"
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
               Begin Threed.SSPanel SSPanel39 
                  Height          =   285
                  Left            =   14550
                  TabIndex        =   79
                  Top             =   60
                  Width           =   1245
                  _Version        =   65536
                  _ExtentX        =   2196
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
                  Begin VB.CheckBox chkSelectAut 
                     BackColor       =   &H00004000&
                     Caption         =   "Check1"
                     Height          =   255
                     Left            =   930
                     TabIndex        =   81
                     Top             =   20
                     Width           =   255
                  End
               End
               Begin Threed.SSPanel SSPanel40 
                  Height          =   285
                  Left            =   2565
                  TabIndex        =   80
                  Top             =   60
                  Width           =   3690
                  _Version        =   65536
                  _ExtentX        =   6509
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Trabajador"
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
               Begin Threed.SSPanel SSPanel41 
                  Height          =   285
                  Left            =   13440
                  TabIndex        =   83
                  Top             =   60
                  Width           =   1125
                  _Version        =   65536
                  _ExtentX        =   1984
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Situación"
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
               Begin Threed.SSPanel SSPanel28 
                  Height          =   285
                  Left            =   6240
                  TabIndex        =   73
                  Top             =   60
                  Width           =   1260
                  _Version        =   65536
                  _ExtentX        =   2222
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Tipo Operación"
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
            Begin Threed.SSPanel SSPanel42 
               Height          =   645
               Left            =   -74970
               TabIndex        =   84
               Top             =   1080
               Width           =   15885
               _Version        =   65536
               _ExtentX        =   28019
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
               Begin VB.ComboBox cmb_BusPor 
                  Height          =   315
                  Left            =   1170
                  Style           =   2  'Dropdown List
                  TabIndex        =   85
                  Top             =   210
                  Width           =   2175
               End
               Begin VB.ComboBox cmb_SitAut 
                  Height          =   315
                  Left            =   9060
                  Style           =   2  'Dropdown List
                  TabIndex        =   88
                  Top             =   210
                  Width           =   2830
               End
               Begin EditLib.fpDateTime ipp_FecIniAut 
                  Height          =   315
                  Left            =   5115
                  TabIndex        =   86
                  Top             =   210
                  Width           =   1365
                  _Version        =   196608
                  _ExtentX        =   2408
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
                  Text            =   "28/09/2004"
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
               Begin EditLib.fpDateTime ipp_FecFinAut 
                  Height          =   315
                  Left            =   6510
                  TabIndex        =   87
                  Top             =   210
                  Width           =   1365
                  _Version        =   196608
                  _ExtentX        =   2408
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
                  Text            =   "28/09/2004"
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
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Buscar por:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   93
                  Top             =   270
                  Width           =   810
               End
               Begin VB.Label lbl_BusPor 
                  AutoSize        =   -1  'True
                  Caption         =   "Fecha Operación:"
                  Height          =   195
                  Left            =   3750
                  TabIndex        =   90
                  Top             =   270
                  Width           =   1275
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Situación:"
                  Height          =   195
                  Left            =   8250
                  TabIndex        =   89
                  Top             =   270
                  Width           =   705
               End
            End
         End
      End
   End
   Begin VB.Menu MnuPopUp 
      Caption         =   "MnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu smnu 
         Caption         =   "Exportar a Excel Listado"
         Index           =   0
      End
      Begin VB.Menu smnu 
         Caption         =   "Exportar a Excel Detallado"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frm_Ctb_GesPer_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type arr_PerVac
   perVac_Item    As Long
   perVac_Situac  As String
   perVac_FecIni  As String
   perVac_FecFin  As String
   perVac_DiaAcu  As Long
   perVac_DiaGoz  As Long
   perVac_DiaDis  As Long
End Type

Dim l_arr_PerVac()  As arr_PerVac
Dim l_arr_VacSol()  As arr_PerVac
Dim l_arr_Empres()  As moddat_tpo_Genera
Dim l_arr_Sucurs()  As moddat_tpo_Genera
Dim l_int_PerMes    As Integer
Dim l_int_PerAno    As Integer

Private Sub chkSelectAut_Click()
Dim r_int_Fila As Integer
   
   If grd_ListAut.Rows > 0 Then
      If chkSelectAut.Value = 0 Then
         For r_int_Fila = 0 To grd_ListAut.Rows - 1
             grd_ListAut.TextMatrix(r_int_Fila, 9) = ""
         Next
      End If
      If chkSelectAut.Value = 1 Then
         For r_int_Fila = 0 To grd_ListAut.Rows - 1
             grd_ListAut.TextMatrix(r_int_Fila, 9) = "X"
         Next
      End If
   Call gs_RefrescaGrid(grd_ListAut)
   End If
End Sub

Private Sub chkSelectPag_Click()
Dim r_Fila As Integer
   
   If grd_ListPag.Rows > 0 Then
      If chkSelectPag.Value = 0 Then
         For r_Fila = 0 To grd_ListPag.Rows - 1
             If UCase(grd_ListPag.TextMatrix(r_Fila, 7)) = "NO" Then
                grd_ListPag.TextMatrix(r_Fila, 10) = ""
             End If
         Next r_Fila
      End If
      If chkSelectPag.Value = 1 Then
         For r_Fila = 0 To grd_ListPag.Rows - 1
             If UCase(grd_ListPag.TextMatrix(r_Fila, 7)) = "NO" Then
                grd_ListPag.TextMatrix(r_Fila, 10) = "X"
             End If
         Next r_Fila
      End If
      Call gs_RefrescaGrid(grd_ListPag)
   End If
End Sub

Private Sub cmb_BusPag_Click()
    If (cmb_BusPag.ListIndex = 0 Or cmb_BusPag.ListIndex = -1) Then
        txt_BusPag.Enabled = False
        Call gs_SetFocus(cmd_BusPag)
    Else
        txt_BusPag.Enabled = True
    End If
    txt_BusPag.Text = ""
End Sub

Private Sub cmb_BusPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_BusPag.Enabled = False) Then
          Call gs_SetFocus(cmd_BusPag)
      Else
          Call gs_SetFocus(txt_BusPag)
      End If
   End If
End Sub

Private Sub cmb_BusPor_Click()
 If cmb_BusPor.ListIndex > -1 Then
    If cmb_BusPor.ListIndex = 0 Then
       lbl_BusPor.Caption = "Fecha Operación:"
       ipp_FecIniAut.Text = DateAdd("m", -5, moddat_g_str_FecSis)
       ipp_FecFinAut.Text = DateAdd("m", 12, moddat_g_str_FecSis)
    ElseIf cmb_BusPor.ListIndex = 1 Then
       lbl_BusPor.Caption = "Fecha Goce Vac.:"
       ipp_FecIniAut.Text = moddat_g_str_FecSis
       ipp_FecFinAut.Text = ff_Ultimo_Dia_Mes(Format(moddat_g_str_FecSis, "mm"), Format(moddat_g_str_FecSis, "yyyy")) & "/" & Format(moddat_g_str_FecSis, "mm") & "/" & Format(moddat_g_str_FecSis, "yyyy")
    End If
 End If
End Sub

Private Sub cmb_BusPor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIniAut)
   End If
End Sub

Private Sub cmb_BusVac_Click()
    If (cmb_BusVac.ListIndex = 0 Or cmb_BusVac.ListIndex = -1) Then
        txt_BusVac.Enabled = False
        Call gs_SetFocus(cmd_BusVac)
    Else
        txt_BusVac.Enabled = True
    End If
    txt_BusVac.Text = ""
End Sub

Private Sub cmb_BusVac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_BusVac.Enabled = False) Then
          Call gs_SetFocus(cmd_BusVac)
      Else
          Call gs_SetFocus(txt_BusVac)
      End If
   End If
End Sub

Private Sub cmb_SitAut_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusAut)
   End If
End Sub

Private Sub cmd_AgregaPag_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_TipRec = 1 'GESTION DE PAGOS
   frm_Ctb_GesPer_02.Show 1
End Sub

Private Sub cmd_AgregaVac_Click()
   moddat_g_int_FlgGrb = 1
   moddat_g_int_TipRec = 2 'GESTION DE VACACIONES
   frm_Ctb_GesPer_02.Show 1
End Sub

Private Sub cmd_AprAut_Click()
Dim r_int_Fila    As Integer
Dim r_str_CodGrb  As String
Dim r_bol_Estado  As Boolean
Dim r_str_Parame  As String
Dim r_rst_Genera  As ADODB.Recordset
Dim r_rst_Princi  As ADODB.Recordset
Dim r_int_Item    As Integer

   If grd_ListAut.Rows = 0 Then
      Exit Sub
   End If
   
   r_bol_Estado = False
   For r_int_Fila = 0 To grd_ListAut.Rows - 1
       If Trim(grd_ListAut.TextMatrix(r_int_Fila, 9)) = "X" Then
          If CLng(grd_ListAut.TextMatrix(grd_ListAut.Row, 10)) <> 1 Then
             MsgBox "Solo se aceptan registros con situación pendiente.", vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          r_bol_Estado = True
       End If
   Next
            
   If r_bol_Estado = False Then
      MsgBox "No hay ninguna fila seleccionada.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   Call gs_RefrescaGrid(grd_ListAut)
   If MsgBox("¿Seguro que desea aprobar lo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   For r_int_Fila = 0 To grd_ListAut.Rows - 1
       If Trim(grd_ListAut.TextMatrix(r_int_Fila, 9)) = "X" Then

          r_str_Parame = ""
          r_str_Parame = r_str_Parame & " USP_CNTBL_GESPER_BORRAR ( "
          r_str_Parame = r_str_Parame & "'" & CLng(grd_ListAut.TextMatrix(r_int_Fila, 0)) & "', " 'GESPER_CODGES
          r_str_Parame = r_str_Parame & "2, " 'TIPO TABLA
          r_str_Parame = r_str_Parame & "2, " 'ESTADO APROBADO
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
          r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
      
          If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
          If r_rst_Genera!RESUL = 1 Then
             r_str_CodGrb = r_str_CodGrb & " - " & CStr(grd_ListAut.TextMatrix(r_int_Fila, 0)) 'COMAUT_CODOPE
          
             ReDim l_arr_PerVac(0)
             Call fs_CalPerido_Vac_02(gf_FormatoFecha(r_rst_Genera!FECING), r_rst_Genera!DIAS_VCTO, grd_ListAut.TextMatrix(r_int_Fila, 11), _
                                      grd_ListAut.TextMatrix(r_int_Fila, 12), l_arr_PerVac, CLng(grd_ListAut.TextMatrix(r_int_Fila, 6)))
             
             For r_int_Item = 1 To UBound(l_arr_PerVac)
                  r_str_Parame = ""
                  r_str_Parame = r_str_Parame & " USP_CNTBL_GESPER ( "
                  r_str_Parame = r_str_Parame & CLng(grd_ListAut.TextMatrix(r_int_Fila, 0)) & ", " 'GESPER_CODGES
                  r_str_Parame = r_str_Parame & grd_ListAut.TextMatrix(r_int_Fila, 11) & ", " 'GESPER_TIPDOC
                  r_str_Parame = r_str_Parame & "'" & Trim(grd_ListAut.TextMatrix(r_int_Fila, 12)) & "', " 'GESPER_NUMDOC
                  r_str_Parame = r_str_Parame & Format(date, "yyyymmdd") & ", " 'GESPER_FECOPE
                  r_str_Parame = r_str_Parame & "NULL, " 'TIPO CAMBIO
                  r_str_Parame = r_str_Parame & grd_ListAut.TextMatrix(r_int_Fila, 13) & ", " 'GESPER_TIPOPE
                  r_str_Parame = r_str_Parame & "NULL, " 'TIPO MONEDA
                  r_str_Parame = r_str_Parame & l_arr_PerVac(r_int_Item).perVac_DiaGoz & ", " 'GESPER_IMPORT
                  r_str_Parame = r_str_Parame & "Null , " 'GESPER_CODBNC
                  r_str_Parame = r_str_Parame & "'', " 'GESPER_CTACRR
                  r_str_Parame = r_str_Parame & "4 , " 'GESPER_TIPTAB - HITORICO
                  r_str_Parame = r_str_Parame & Format(l_arr_PerVac(r_int_Item).perVac_FecIni, "yyyymmdd") & ", " 'GESPER_FECHA1
                  r_str_Parame = r_str_Parame & Format(l_arr_PerVac(r_int_Item).perVac_FecFin, "yyyymmdd") & ", " 'GESPER_FECHA2
                  r_str_Parame = r_str_Parame & "'', " 'GESPER_DESCRI
                  r_str_Parame = r_str_Parame & "NULL," 'GESPER_DIAVEN
                  r_str_Parame = r_str_Parame & "NULL," 'GESPER_DIAVIG
                  r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
                  r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
                  r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
                  r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "') "
                  r_str_Parame = r_str_Parame & "1) " 'as_insupd
            
                  If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
                     Exit Sub
                  End If
             Next
          End If
       End If
   Next
   
   MsgBox "Registros aprobados correctamente." & vbCrLf & "Codigos :" & r_str_CodGrb, vbInformation, modgen_g_str_NomPlt
   Screen.MousePointer = 0
   Call fs_BuscarAut
   Call gs_SetFocus(grd_ListAut)
End Sub

Private Sub cmd_EditVac_Click()
   If grd_ListVac.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_str_CodGen = ""
   moddat_g_int_TipDoc = CStr(grd_ListVac.TextMatrix(grd_ListVac.Row, 10))
   moddat_g_str_NumDoc = CStr(grd_ListVac.TextMatrix(grd_ListVac.Row, 11))

   frm_Ctb_GesPer_04.Show 1
End Sub

Private Sub cmd_ExpAut_Click()
   If grd_ListAut.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc_Aut
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpVac_Click()
   Me.PopupMenu MnuPopUp
End Sub

Private Sub cmd_LimAut_Click()
   cmb_SitAut.ListIndex = 0
   Call gs_LimpiaGrid(grd_ListAut)
   cmb_BusPor.Enabled = True
   ipp_FecIniAut.Enabled = True
   ipp_FecFinAut.Enabled = True
   cmb_SitAut.Enabled = True
   cmb_BusPor.ListIndex = 0
   Call gs_SetFocus(cmb_BusPor)
End Sub

Private Sub cmd_Reversa_Click()
Dim r_int_Fila    As Integer
Dim r_str_CodGrb  As String
Dim r_bol_Estado  As Boolean
Dim r_str_Parame  As String
Dim r_rst_Genera  As ADODB.Recordset

   If grd_ListAut.Rows = 0 Then
      Exit Sub
   End If
   
   r_bol_Estado = False
   For r_int_Fila = 0 To grd_ListAut.Rows - 1
       If Trim(grd_ListAut.TextMatrix(r_int_Fila, 9)) = "X" Then
          If CLng(grd_ListAut.TextMatrix(r_int_Fila, 10)) = 0 Or CLng(grd_ListAut.TextMatrix(r_int_Fila, 10)) = 1 Then
             MsgBox "Solo se puede reversar los registro con situación aprobado o rechazado.", vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          r_bol_Estado = True
       End If
   Next
   
   If r_bol_Estado = False Then
      MsgBox "No hay ninguna fila seleccionada.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   Call gs_RefrescaGrid(grd_ListAut)
   If MsgBox("¿Seguro que desea dar reversa lo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   For r_int_Fila = 0 To grd_ListAut.Rows - 1
       If Trim(grd_ListAut.TextMatrix(r_int_Fila, 9)) = "X" Then
       
          r_str_Parame = ""
          r_str_Parame = r_str_Parame & " USP_CNTBL_GESPER_BORRAR ( "
          r_str_Parame = r_str_Parame & "'" & CLng(grd_ListAut.TextMatrix(r_int_Fila, 0)) & "', " 'GESPER_CODGES
          r_str_Parame = r_str_Parame & "2, " 'TIPO TABLA
          r_str_Parame = r_str_Parame & "-1, " 'ESTADO REVERTIR
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
          r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
      
          If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
          If r_rst_Genera!RESUL = 1 Then
             r_str_CodGrb = r_str_CodGrb & " - " & CStr(grd_ListAut.TextMatrix(r_int_Fila, 0)) 'COMAUT_CODOPE
          End If
       End If
   Next
   
   MsgBox "Registros reversados correctamente." & vbCrLf & "Codigos :" & r_str_CodGrb, vbInformation, modgen_g_str_NomPlt
   Screen.MousePointer = 0
   Call fs_BuscarAut
   Call gs_SetFocus(grd_ListAut)
End Sub

Private Sub cmd_RhzAut_Click()
Dim r_int_Fila    As Integer
Dim r_str_CodGrb  As String
Dim r_bol_Estado  As Boolean
Dim r_str_Parame  As String
Dim r_rst_Genera  As ADODB.Recordset

   If grd_ListAut.Rows = 0 Then
      Exit Sub
   End If
   
   r_bol_Estado = False
   For r_int_Fila = 0 To grd_ListAut.Rows - 1
       If Trim(grd_ListAut.TextMatrix(r_int_Fila, 9)) = "X" Then
          If CLng(grd_ListAut.TextMatrix(grd_ListAut.Row, 10)) <> 1 Then
             MsgBox "Solo se aceptan registros con situación pendiente.", vbExclamation, modgen_g_str_NomPlt
             Exit Sub
          End If
          r_bol_Estado = True
       End If
   Next
   
   If r_bol_Estado = False Then
      MsgBox "No hay ninguna fila seleccionada.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   Call gs_RefrescaGrid(grd_ListAut)
   If MsgBox("¿Seguro que desea rechazar lo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   For r_int_Fila = 0 To grd_ListAut.Rows - 1
       If Trim(grd_ListAut.TextMatrix(r_int_Fila, 9)) = "X" Then

          r_str_Parame = ""
          r_str_Parame = r_str_Parame & " USP_CNTBL_GESPER_BORRAR ( "
          r_str_Parame = r_str_Parame & "'" & CLng(grd_ListAut.TextMatrix(r_int_Fila, 0)) & "', " 'GESPER_CODGES
          r_str_Parame = r_str_Parame & "2, " 'TIPO TABLA
          r_str_Parame = r_str_Parame & "3, " 'ESTADO RECHAZADO
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
          r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
          r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
          If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
          If r_rst_Genera!RESUL = 1 Then
             r_str_CodGrb = r_str_CodGrb & " - " & CStr(grd_ListAut.TextMatrix(r_int_Fila, 0)) 'COMAUT_CODOPE
          End If
       End If
   Next
   
   MsgBox "Registros rechazados correctamente." & vbCrLf & "Codigos :" & r_str_CodGrb, vbInformation, modgen_g_str_NomPlt
   Screen.MousePointer = 0
   Call fs_BuscarAut
   Call gs_SetFocus(grd_ListAut)
End Sub

Private Sub cmd_BorrarPag_Click()
Dim r_str_Parame  As String
Dim r_rst_Princi  As ADODB.Recordset
Dim r_rst_Genera  As ADODB.Recordset

   moddat_g_str_Codigo = ""
   
   If grd_ListPag.Rows = 0 Then
      Exit Sub
   End If
   moddat_g_str_Codigo = CLng(grd_ListPag.TextMatrix(grd_ListPag.Row, 0))
   
   If Trim(grd_ListPag.TextMatrix(grd_ListPag.Row, 7)) = "SI" Then
      '--procesado por Compensasion
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT NVL((SELECT COMAUT_CODEST FROM CNTBL_COMAUT A  "
      r_str_Parame = r_str_Parame & "              Where A.COMAUT_SITUAC = 1  "
      r_str_Parame = r_str_Parame & "                AND A.COMAUT_CODEST IN (1,2,4,5)  "
      r_str_Parame = r_str_Parame & "                AND A.COMAUT_CODOPE = " & CLng(moddat_g_str_Codigo) & ")  "
      r_str_Parame = r_str_Parame & "           ,0) AS CODEST  "
      r_str_Parame = r_str_Parame & "   FROM DUAL  "
    
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
         Screen.MousePointer = 0
         Exit Sub
      End If

      If r_rst_Princi.BOF And r_rst_Princi.EOF Then 'ningún registro
         r_rst_Princi.Close
         Set r_rst_Princi = Nothing
         Screen.MousePointer = 0
         Exit Sub
      End If
   
      r_rst_Princi.MoveFirst
      If r_rst_Princi!CODEST <> 0 Then
         Select Case r_rst_Princi!CODEST
                Case 1: MsgBox "El registro se encuentra como pendiente en modulo de compensación, no se puede eliminar.", vbExclamation, modgen_g_str_NomPlt
                Case 2: MsgBox "El registro se encuentra como aprobado en modulo de compensación, no se puede eliminar.", vbExclamation, modgen_g_str_NomPlt
                Case 4: MsgBox "El registro se encuentra como aplicado en modulo de compensación, no se puede eliminar.", vbExclamation, modgen_g_str_NomPlt
                Case 5: MsgBox "El registro se encuentra como pagado en modulo de compensación, no se puede eliminar.", vbExclamation, modgen_g_str_NomPlt
         End Select
         Exit Sub
      End If
      '----------------------------------------
      If MsgBox("¿Seguro que desea eliminar el registro seleccionado?" & vbCrLf & _
                "Recuerde que debe eliminar el asiento contable manual.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   Else
      Call gs_RefrescaGrid(grd_ListPag)
      If MsgBox("¿Seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   End If
   
   Call gs_RefrescaGrid(grd_ListPag)
   
   Screen.MousePointer = 11
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " USP_CNTBL_GESPER_BORRAR ( "
   r_str_Parame = r_str_Parame & "'" & CLng(moddat_g_str_Codigo) & "', " 'GESPER_CODGES
   r_str_Parame = r_str_Parame & "1, " 'TIPO TABLA
   r_str_Parame = r_str_Parame & "0, " 'ESTADO
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
   r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 2) Then
      MsgBox "No se pudo completar la eliminación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   Else
      MsgBox "El registro se elimino, recuerde que debe eliminar el asiento contable manual.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
   
   Call fs_BuscarPag
   Call gs_SetFocus(grd_ListPag)
End Sub

Private Sub cmd_BusAut_Click()
   Call fs_BuscarAut
   cmb_BusPor.Enabled = False
   ipp_FecIniAut.Enabled = False
   ipp_FecFinAut.Enabled = False
   cmb_SitAut.Enabled = False
   Call gs_SetFocus(grd_ListAut)
End Sub

Private Sub cmd_BusPag_Click()
   Call fs_BuscarPag
   cmb_EmpPag.Enabled = False
   cmb_SucPag.Enabled = False
   ipp_FecIniPag.Enabled = False
   ipp_FecFinPag.Enabled = False
End Sub

Private Sub cmd_BusVac_Click()
   Call fs_BuscarVac
End Sub

Private Sub cmd_ConsulAut_Click()
   If grd_ListAut.Rows = 0 Then
      Exit Sub
   End If
   
   grd_ListAut.Col = 0
   moddat_g_str_Codigo = CLng(grd_ListAut.Text)

   grd_ListAut.Col = 11
   moddat_g_int_TipDoc = CStr(grd_ListAut.Text)
   
   grd_ListAut.Col = 12
   moddat_g_str_NumDoc = CStr(grd_ListAut.Text)
      
   moddat_g_int_FlgGrb = 0 'consultar
   
   Call gs_RefrescaGrid(grd_ListAut)
   frm_Ctb_GesPer_02.Show 1
   
   Call gs_SetFocus(grd_ListAut)
End Sub

Private Sub cmd_ConsulPag_Click()
   moddat_g_int_FlgGrb = 0
   moddat_g_str_Codigo = ""
   
   If grd_ListPag.Rows = 0 Then
      Exit Sub
   End If
   Call gs_RefrescaGrid(grd_ListPag)
   
   moddat_g_str_Codigo = grd_ListPag.TextMatrix(grd_ListPag.Row, 0)
   Call gs_RefrescaGrid(grd_ListPag)
   
   moddat_g_int_TipRec = 1 'GESTION DE PAGOS
   moddat_g_int_FlgGrb = 0
   frm_Ctb_GesPer_02.Show 1
   
   Call gs_SetFocus(grd_ListPag)
End Sub

Private Sub cmd_DetVac_Click()
moddat_g_int_TipDoc = 0
moddat_g_str_NumDoc = ""
moddat_g_str_NomCli = ""
moddat_g_str_FecIng = ""
moddat_g_str_CodGen = ""

   If grd_ListVac.Rows = 0 Then
      Exit Sub
   End If
   Call gs_RefrescaGrid(grd_ListVac)
   
   moddat_g_int_TipRec = 2 'GESTION DE VACACIONES
   
   grd_ListVac.Col = 0
   moddat_g_str_CodGen = CStr(grd_ListVac.Text)
   
   grd_ListVac.Col = 2
   moddat_g_str_NomCli = CStr(grd_ListVac.Text)
   
   grd_ListVac.Col = 4
   moddat_g_str_FecIng = CStr(grd_ListVac.Text)
   
   grd_ListVac.Col = 10
   moddat_g_int_TipDoc = CStr(grd_ListVac.Text)
   
   grd_ListVac.Col = 11
   moddat_g_str_NumDoc = CStr(grd_ListVac.Text)

   Call gs_RefrescaGrid(grd_ListVac)
   
   frm_Ctb_GesPer_03.Show 1
   
   Call gs_SetFocus(grd_ListVac)
End Sub

Private Sub cmd_ExpPag_Click()
   If grd_ListPag.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExcPag
   Screen.MousePointer = 0
End Sub

Private Sub cmd_GenPag_Click()
Dim r_int_Contad        As Integer
Dim r_bol_Estado        As Boolean

   r_bol_Estado = False
   For r_int_Contad = 0 To grd_ListPag.Rows - 1
       If grd_ListPag.TextMatrix(r_int_Contad, 7) = "NO" Then
          If grd_ListPag.TextMatrix(r_int_Contad, 10) = "X" Then
             r_bol_Estado = True
             Exit For
          End If
       End If
   Next
   
   If r_bol_Estado = False Then
      MsgBox "No se han seleccionados registros para generar asientos automáticos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If

   'confirma
   If MsgBox("¿Está seguro de generar los asientos contables?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GeneraAsiPag
   Call cmd_BusPag_Click
   Screen.MousePointer = 0
End Sub

Private Sub cmd_LimPag_Click()
   Call fs_LimPag
   cmb_EmpPag.Enabled = True
   cmb_SucPag.Enabled = True
   ipp_FecIniPag.Enabled = True
   ipp_FecFinPag.Enabled = True
   Call gs_SetFocus(cmb_EmpPag)
End Sub

Private Sub cmd_LimVac_Click()
   Call fs_LimPag
   Call gs_SetFocus(grd_ListVac)
End Sub

Private Sub cmd_SalirAut_Click()
   Unload Me
End Sub

Private Sub cmd_SalirPag_Click()
   Unload Me
End Sub

Private Sub cmd_SalirVac_Click()
   Unload Me
End Sub


Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_LimPag
   
   tab_GesPer.TabEnabled(0) = False
   tab_GesPer.TabEnabled(1) = False
   tab_GesPer.TabEnabled(2) = False
   
   If moddat_g_int_TipRec = 1 Then
      'GESTION DE PAGOS
      tab_GesPer.TabEnabled(0) = True
      tab_GesPer.Tab = 0
   ElseIf moddat_g_int_TipRec = 2 Then
       'GESTION DE VACAIONES
      tab_GesPer.TabEnabled(1) = True
      tab_GesPer.Tab = 1
      cmd_EditVac.Enabled = False
      cmd_ExpVac.Enabled = False
      Call fs_BuscarVac
   ElseIf moddat_g_int_TipRec = 3 Then
      'AUTORIZACION DE VACACIONES
      tab_GesPer.TabEnabled(2) = True
      tab_GesPer.Tab = 2
      Call cmd_BusAut_Click
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   Call moddat_gs_Carga_EmpGrp(cmb_EmpPag, l_arr_Empres)
   
   cmb_BusPag.Clear
   cmb_BusPag.AddItem "NINGUNA"
   cmb_BusPag.AddItem "NRO DOCUMENTO"
   cmb_BusPag.AddItem "TIPO OPERACION"
   cmb_BusPag.AddItem "TRABAJADOR"
   
   cmb_BusVac.Clear
   cmb_BusVac.AddItem "NINGUNA"
   cmb_BusVac.AddItem "CODIGO PLANILLA"
   cmb_BusVac.AddItem "NRO DOCUMENTO"
   cmb_BusVac.AddItem "TRABAJADOR"
   
   cmb_BusPor.Clear
   cmb_BusPor.AddItem "FECHA OPERACION"
   cmb_BusPor.AddItem "FECHA GOCE VAC."
   cmb_BusPor.ListIndex = 0
   
   cmb_SitAut.Clear
   cmb_SitAut.AddItem "PENDIENTES"
   cmb_SitAut.ItemData(cmb_SitAut.NewIndex) = 1
   cmb_SitAut.AddItem "APROBADOS"
   cmb_SitAut.ItemData(cmb_SitAut.NewIndex) = 2
   cmb_SitAut.AddItem "RECHAZADOS"
   cmb_SitAut.ItemData(cmb_SitAut.NewIndex) = 3
   cmb_SitAut.AddItem "<<TODOS>>"
   cmb_SitAut.ItemData(cmb_SitAut.NewIndex) = 0
   cmb_SitAut.ListIndex = -1
   Call gs_BuscarCombo_Item(cmb_SitAut, 1)
   
   ipp_FecIniAut.Text = DateAdd("m", -5, moddat_g_str_FecSis)
   ipp_FecFinAut.Text = DateAdd("m", 12, moddat_g_str_FecSis)
   
   'GESTION DE PAGOS
   grd_ListPag.ColWidth(0) = 1140 'CODIGO
   grd_ListPag.ColWidth(1) = 1250 'NRO DOCUMENTO
   grd_ListPag.ColWidth(2) = 2400 'TIPO OPERACION
   grd_ListPag.ColWidth(3) = 3270 'TRABAJADOR
   grd_ListPag.ColWidth(4) = 1060 'FECHA
   grd_ListPag.ColWidth(5) = 870 'MONEDA
   grd_ListPag.ColWidth(6) = 1140 'IMPORTE
   grd_ListPag.ColWidth(7) = 1050 'CONTABILIDAD
   grd_ListPag.ColWidth(8) = 1060 'FECHA PAGO
   grd_ListPag.ColWidth(9) = 1040 'CODIGO PAGO
   
   grd_ListPag.ColWidth(10) = 1170 'SELECCION CTA
   grd_ListPag.ColWidth(11) = 0 'SELECCION TXT
   grd_ListPag.ColWidth(12) = 0 'TIPO DE OPERACION
   grd_ListPag.ColWidth(13) = 0 'TIPO DE MONEDA
   grd_ListPag.ColWidth(14) = 0 'NRO DOCUMENTO
   grd_ListPag.ColWidth(15) = 0 'TIPO DOCUMENTO
   grd_ListPag.ColWidth(16) = 0 'TIPO CUENTA
   grd_ListPag.ColWidth(17) = 0 'CTA CORRIENTE TXT
   grd_ListPag.ColWidth(18) = 0 'CODIGO BANCO
   grd_ListPag.ColWidth(19) = 0 'CUENTA CORRIENTE TABLA
   
   grd_ListPag.ColAlignment(0) = flexAlignCenterCenter
   grd_ListPag.ColAlignment(1) = flexAlignCenterCenter
   grd_ListPag.ColAlignment(2) = flexAlignLeftCenter
   grd_ListPag.ColAlignment(3) = flexAlignLeftCenter
   grd_ListPag.ColAlignment(4) = flexAlignCenterCenter
   grd_ListPag.ColAlignment(5) = flexAlignLeftCenter
   grd_ListPag.ColAlignment(6) = flexAlignRightCenter
   grd_ListPag.ColAlignment(7) = flexAlignCenterCenter
   grd_ListPag.ColAlignment(8) = flexAlignCenterCenter
   grd_ListPag.ColAlignment(9) = flexAlignCenterCenter
   grd_ListPag.ColAlignment(10) = flexAlignCenterCenter
      
   'GESTION DE VACACIONES
   grd_ListVac.ColWidth(0) = 1220 'CODIGO PLANILLA
   grd_ListVac.ColWidth(1) = 1740 'NRO DOCUMENTO
   grd_ListVac.ColWidth(2) = 4710 'TRABAJADOR
   grd_ListVac.ColWidth(3) = 1400 'ESTADO
   grd_ListVac.ColWidth(4) = 1560 'FECHA INGRESO
   grd_ListVac.ColWidth(5) = 0 'GANADOS(DIAS)
   grd_ListVac.ColWidth(6) = 1470 'VENCIDOS(DIAS)
   grd_ListVac.ColWidth(7) = 1570 'GOZADOS(DIAS)
   grd_ListVac.ColWidth(8) = 0 'SALDO(DIAS)
   grd_ListVac.ColWidth(9) = 1790 'SALDO VENCIDO(DIAS)
   grd_ListVac.ColWidth(10) = 0 'TIPO DOCUMENTO
   grd_ListVac.ColWidth(11) = 0 'NRO DOCUMENTO
   grd_ListVac.ColWidth(12) = 0 'FECHA INGRESO
   grd_ListVac.ColWidth(13) = 0 'CODIGO GESTION
   
   grd_ListVac.ColAlignment(0) = flexAlignCenterCenter
   grd_ListVac.ColAlignment(1) = flexAlignCenterCenter
   grd_ListVac.ColAlignment(2) = flexAlignLeftCenter
   grd_ListVac.ColAlignment(3) = flexAlignCenterCenter
   grd_ListVac.ColAlignment(4) = flexAlignCenterCenter
   grd_ListVac.ColAlignment(5) = flexAlignRightCenter
   grd_ListVac.ColAlignment(6) = flexAlignRightCenter
   grd_ListVac.ColAlignment(7) = flexAlignRightCenter
   grd_ListVac.ColAlignment(8) = flexAlignRightCenter
   grd_ListVac.ColAlignment(9) = flexAlignRightCenter
   
   'GESTION DE AUTORIZACION
   grd_ListAut.ColWidth(0) = 1190 'CODIGO INTERNO
   grd_ListAut.ColWidth(1) = 1320 'FECHA OPERACION
   grd_ListAut.ColWidth(2) = 3680 'TRABAJADOR
   grd_ListAut.ColWidth(3) = 1250 'TIPO OPERACION
   grd_ListAut.ColWidth(4) = 1050 'FECHA DESDE
   grd_ListAut.ColWidth(5) = 1030 'FECHA HASTA
   grd_ListAut.ColWidth(6) = 720 'DIAS SOLICITADOS
   grd_ListAut.ColWidth(7) = 3150 'COMENTARIO
   grd_ListAut.ColWidth(8) = 1100 'SITUACION
   grd_ListAut.ColWidth(9) = 960 'SELECCION
   grd_ListAut.ColWidth(10) = 0 'CODIGO - SITUACION
   grd_ListAut.ColWidth(11) = 0 'TIPO DOCUMENTO
   grd_ListAut.ColWidth(12) = 0 'NRO DOCUMENTO
   grd_ListAut.ColWidth(13) = 0 'CODIGO TIPO OPERACION
   
   grd_ListAut.ColAlignment(0) = flexAlignCenterCenter
   grd_ListAut.ColAlignment(1) = flexAlignCenterCenter
   grd_ListAut.ColAlignment(2) = flexAlignLeftCenter
   grd_ListAut.ColAlignment(3) = flexAlignLeftCenter
   grd_ListAut.ColAlignment(4) = flexAlignCenterCenter
   grd_ListAut.ColAlignment(5) = flexAlignCenterCenter
   grd_ListAut.ColAlignment(6) = flexAlignCenterCenter
   grd_ListAut.ColAlignment(7) = flexAlignLeftCenter
   grd_ListAut.ColAlignment(8) = flexAlignLeftCenter
   grd_ListAut.ColAlignment(9) = flexAlignCenterCenter
End Sub

Private Sub fs_LimPag()
Dim r_str_CadAux As String

   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0
   cmb_EmpPag.ListIndex = 0
   r_str_CadAux = ""
   
   Call moddat_gs_Carga_SucAge(cmb_SucPag, l_arr_Sucurs, l_arr_Empres(cmb_EmpPag.ListIndex + 1).Genera_Codigo)
   
   pnl_PerPag.Caption = moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_EmpPag.ListIndex + 1).Genera_Codigo, 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   r_str_CadAux = DateAdd("m", 1, "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno)
   modctb_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   modctb_str_FecIni = DateAdd("m", -1, modctb_str_FecFin)
   modctb_str_FecIni = "01/" & Format(Month(modctb_str_FecIni), "00") & "/" & Year(modctb_str_FecIni)
   
   ipp_FecIniPag.Text = modctb_str_FecIni
   ipp_FecFinPag.Text = modctb_str_FecFin
   
   cmb_BusPag.ListIndex = 0
   cmb_SucPag.ListIndex = 0
   Call gs_LimpiaGrid(grd_ListPag)
      
   cmb_BusVac.ListIndex = 0
   Call gs_LimpiaGrid(grd_ListVac)
   Call gs_LimpiaGrid(grd_ListAut)
End Sub

Public Sub fs_BuscarPag()
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_str_Cadena     As String
Dim r_str_FecVct     As String
Dim r_str_FecApe     As String
Dim r_int_FecDif     As Integer
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_ListPag)
   r_str_FecIni = Format(ipp_FecIniPag.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFinPag.Text, "yyyymmdd")
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT LPAD(A.GESPER_CODGES,10,'0') AS GESPER_CODGES, A.GESPER_TIPDOC || '-' || A.GESPER_NUMDOC IDPERSONAL,  "
   g_str_Parame = g_str_Parame & "         TRIM(C.PARDES_DESCRI) AS TIPO_OPERACION, TRIM(B.MAEPRV_RAZSOC) AS MAEPRV_RAZSOC,  "
   g_str_Parame = g_str_Parame & "         A.GESPER_FECOPE, TRIM(D.PARDES_DESCRI) AS MONEDA, A.GESPER_IMPORT,  A.GESPER_NUMDOC, A.GESPER_TIPDOC, "
   g_str_Parame = g_str_Parame & "         DECODE(A.GESPER_DATCNT,NULL,'NO','SI') AS CONTABILIZADO, A.GESPER_TIPOPE, A.GESPER_CODMON,  "
   g_str_Parame = g_str_Parame & "         DECODE(B.MAEPRV_CODBNC_MN1,11,'P','I') AS TIPCTA,  "
   '-----------------------TERCER CAMBIO---------------------------------
   g_str_Parame = g_str_Parame & "         GESPER_CODBNC, GESPER_CTACRR AS NUM_CUENTA_TAB,  "
   g_str_Parame = g_str_Parame & "         DECODE(GESPER_CODBNC,11,SUBSTR(TRIM(GESPER_CTACRR),1,8) || '00'|| SUBSTR(TRIM(GESPER_CTACRR),9,10), GESPER_CTACRR) AS NUM_CUENTA_TXT,  "
   '--------------------------------------------------------
   g_str_Parame = g_str_Parame & "         F.COMPAG_FECPAG, F.COMPAG_CODCOM "
   g_str_Parame = g_str_Parame & "    FROM CNTBL_GESPER A  "
   g_str_Parame = g_str_Parame & "   INNER JOIN CNTBL_MAEPRV B ON A.GESPER_TIPDOC = B.MAEPRV_TIPDOC AND A.GESPER_NUMDOC = B.MAEPRV_NUMDOC  "
   g_str_Parame = g_str_Parame & "   INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 130 AND A.GESPER_TIPOPE = C.PARDES_CODITE  "
   g_str_Parame = g_str_Parame & "   INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND A.GESPER_CODMON = D.PARDES_CODITE  "
   g_str_Parame = g_str_Parame & "    LEFT JOIN CNTBL_COMDET E ON E.COMDET_CODOPE = A.GESPER_CODGES AND E.COMDET_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "    LEFT JOIN CNTBL_COMPAG F ON F.COMPAG_CODCOM = E.COMDET_CODCOM AND F.COMPAG_SITUAC = 1 AND F.COMPAG_FLGCTB = 1  "
   g_str_Parame = g_str_Parame & "   WHERE A.GESPER_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "     AND A.GESPER_FECOPE BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   g_str_Parame = g_str_Parame & "     AND A.GESPER_TIPTAB = 1  " 'Gestion de pagos
   
   If (cmb_BusPag.ListIndex = 1) Then 'NRO DOCUMENTO
       If Len(Trim(txt_BusPag.Text)) > 0 Then
          g_str_Parame = g_str_Parame & "   AND TRIM(B.MAEPRV_NUMDOC) LIKE '%" & UCase(Trim(txt_BusPag.Text)) & "%'"
       End If
   ElseIf (cmb_BusPag.ListIndex = 2) Then 'TIPO OPERACION
       If Len(Trim(txt_BusPag.Text)) > 0 Then
          g_str_Parame = g_str_Parame & "   AND TRIM(C.PARDES_DESCRI) LIKE '%" & UCase(Trim(txt_BusPag.Text)) & "%'"
       End If
   ElseIf (cmb_BusPag.ListIndex = 3) Then 'TRABAJADOR
       If Len(Trim(txt_BusPag.Text)) > 0 Then
          g_str_Parame = g_str_Parame & "   AND TRIM(B.MAEPRV_RAZSOC) LIKE '%" & UCase(Trim(txt_BusPag.Text)) & "%'"
       End If
   End If
   g_str_Parame = g_str_Parame & " ORDER BY A.GESPER_CODGES ASC "

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
   
   grd_ListPag.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_ListPag.Rows = grd_ListPag.Rows + 1
      grd_ListPag.Row = grd_ListPag.Rows - 1
      
      grd_ListPag.Col = 0
      grd_ListPag.Text = CStr(g_rst_Princi!GESPER_CODGES)
      
      grd_ListPag.Col = 1
      grd_ListPag.Text = Trim(g_rst_Princi!IDPERSONAL & "")
      
      grd_ListPag.Col = 2
      grd_ListPag.Text = Trim(g_rst_Princi!TIPO_OPERACION & "")
      
      grd_ListPag.Col = 3
      grd_ListPag.Text = Trim(g_rst_Princi!MaePrv_RazSoc & "")
      
      grd_ListPag.Col = 4
      grd_ListPag.Text = gf_FormatoFecha(g_rst_Princi!GESPER_FECOPE)
      
      grd_ListPag.Col = 5
      grd_ListPag.Text = Trim(g_rst_Princi!Moneda)
      
      grd_ListPag.Col = 6
      grd_ListPag.Text = Format(g_rst_Princi!GESPER_IMPORT, "###,###,###,##0.00")
      
      grd_ListPag.Col = 7
      grd_ListPag.Text = Trim(g_rst_Princi!CONTABILIZADO & "")
      
      If Trim(g_rst_Princi!COMPAG_FECPAG & "") <> "" Then
         grd_ListPag.Col = 8
         grd_ListPag.Text = gf_FormatoFecha(g_rst_Princi!COMPAG_FECPAG)
      End If
      If Trim(g_rst_Princi!COMPAG_CODCOM & "") <> "" Then
         grd_ListPag.Col = 9
         grd_ListPag.Text = Format(g_rst_Princi!COMPAG_CODCOM, "00000000")
      End If
               
      grd_ListPag.Col = 12
      grd_ListPag.Text = Trim(g_rst_Princi!GESPER_TIPOPE & "")
      
      grd_ListPag.Col = 13
      grd_ListPag.Text = Trim(g_rst_Princi!GESPER_CODMON & "")
      
      grd_ListPag.Col = 14
      grd_ListPag.Text = Trim(g_rst_Princi!GESPER_NUMDOC & "")
      
      grd_ListPag.Col = 15
      grd_ListPag.Text = g_rst_Princi!GESPER_TIPDOC
      
      grd_ListPag.Col = 16
      grd_ListPag.Text = Trim(g_rst_Princi!TIPCTA & "")
      
      grd_ListPag.Col = 17
      grd_ListPag.Text = Trim(g_rst_Princi!NUM_CUENTA_TXT & "")
      
      grd_ListPag.Col = 18
      grd_ListPag.Text = Trim(g_rst_Princi!GESPER_CODBNC & "") 'MAEPRV_CODBNC_MN1
      
      grd_ListPag.Col = 19
      grd_ListPag.Text = Trim(g_rst_Princi!NUM_CUENTA_TAB & "")
                  
      g_rst_Princi.MoveNext
   Loop
   
   grd_ListPag.Redraw = True
   Call gs_UbiIniGrid(grd_ListPag)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Public Sub fs_SaldoDias(p_TipDoc As Integer, p_NumDoc As String, ByRef p_Vencido As Integer, ByRef p_Gozados As Integer)
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset

   p_Vencido = 0
   p_Gozados = 0
   r_str_Parame = ""
   'r_str_Parame = r_str_Parame & " SELECT DECODE(A.MAEPRV_FECING,'',0, ROUND((((TO_DATE(TO_CHAR(sysdate,'yyyymmdd'),'YYYYMMDD') - TO_DATE((A.MAEPRV_FECING), 'YYYYMMDD'))*30)/360),0))) AS VAC_GANADAS,   "
   r_str_Parame = r_str_Parame & "  SELECT TRUNC(DECODE(A.MAEPRV_FECING,'',0, ABS((((TO_DATE(TO_CHAR(sysdate,'yyyymmdd'),'YYYYMMDD') - TO_DATE((A.MAEPRV_FECING), 'YYYYMMDD'))*30)/365))),0) AS VAC_GANADAS,  "
   
   r_str_Parame = r_str_Parame & "           trunc(months_between(SYSDATE, to_date(A.MAEPRV_FECING,'YYYYMMDD'))/12) * 30 AS DIAS_VENCIDOS, "
   r_str_Parame = r_str_Parame & "           (NVL(C.GESPER_DIAGOZ,0) + ROUND(NVL(B.DIAS_GOZADOS,0),0)) AS DIAS_GOZADOS  "
   r_str_Parame = r_str_Parame & "      FROM CNTBL_MAEPRV A  "
   r_str_Parame = r_str_Parame & "      LEFT JOIN CNTBL_GESPER C ON C.GESPER_TIPDOC = A.MAEPRV_TIPDOC AND C.GESPER_NUMDOC = A.MAEPRV_NUMDOC "
   r_str_Parame = r_str_Parame & "            AND C.GESPER_SITUAC = 1 AND C.GESPER_TIPTAB = 3 " '--MAESTRO
   r_str_Parame = r_str_Parame & "      LEFT JOIN (SELECT SUM(NVL(H.GESPER_IMPORT,0)) AS DIAS_GOZADOS , H.GESPER_TIPDOC, H.GESPER_NUMDOC  "
   r_str_Parame = r_str_Parame & "                   FROM CNTBL_GESPER H  "
   r_str_Parame = r_str_Parame & "                  Where H.GESPER_TIPTAB = 2  "
   r_str_Parame = r_str_Parame & "                    AND H.GESPER_SITUAC = 2  "
   r_str_Parame = r_str_Parame & "                  GROUP BY H.GESPER_TIPDOC, H.GESPER_NUMDOC) B  "
   r_str_Parame = r_str_Parame & "        ON A.MAEPRV_TIPDOC = B.GESPER_TIPDOC AND A.MAEPRV_NUMDOC = B.GESPER_NUMDOC  "
   r_str_Parame = r_str_Parame & "     WHERE A.MAEPRV_TIPPER = 2  " '--PERSONAL INTERNO"
   r_str_Parame = r_str_Parame & "       AND A.MAEPRV_SITUAC = 1  "
   r_str_Parame = r_str_Parame & "       AND A.MAEPRV_TIPDOC = " & p_TipDoc
   r_str_Parame = r_str_Parame & "       AND A.MAEPRV_NUMDOC = '" & p_NumDoc & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      r_rst_Princi.MoveFirst
       p_Vencido = CInt(r_rst_Princi!DIAS_VENCIDOS) 'CInt(r_rst_Princi!VAC_GANADAS)
       p_Gozados = CInt(r_rst_Princi!DIAS_GOZADOS)
   End If
       
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Public Function fs_UserEjecutivo(p_CodUsu As String, p_CodEje As String) As String
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
    
   fs_UserEjecutivo = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT A.EJETIP_CODEJE FROM CRE_EJETIP A, CRE_EJECMC B "
   r_str_Parame = r_str_Parame & "  WHERE A.EJETIP_CODEJE = B.EJECMC_CODEJE "
   r_str_Parame = r_str_Parame & "    AND A.EJETIP_TIPEJE = " & p_CodEje
   r_str_Parame = r_str_Parame & "    AND A.EJETIP_TIPEJE = " & p_CodEje
   r_str_Parame = r_str_Parame & "    AND B.EJECMC_SITUAC = 1 "
   r_str_Parame = r_str_Parame & "    AND UPPER(TRIM(A.EJETIP_CODEJE)) = '" & UCase(Trim(p_CodUsu)) & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Exit Function
   End If
   
   r_rst_Princi.MoveFirst
   fs_UserEjecutivo = Trim(r_rst_Princi!EJETIP_CODEJE & "")
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
End Function

Public Sub fs_BuscarVac()
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_str_Cadena     As String
Dim r_str_FecVct     As String
Dim r_str_FecApe     As String
Dim r_int_FecDif     As Integer
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
   
Dim r_str_PerAct     As String
Dim r_str_PerAnt     As String
Dim r_str_PerAux     As String

   r_str_PerAux = "01/" & Format(moddat_g_str_FecSis, "mm") & "/" & Format(moddat_g_str_FecSis, "yyyy")
   r_str_PerAct = Format(r_str_PerAux, "yyyymm")
   r_str_PerAnt = Format(DateAdd("d", -1, r_str_PerAux), "yyyymm")
   
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_ListVac)
   '-- =(HOY()-D76)*30/360
   
   r_str_Parame = ""
   'r_str_Parame = r_str_Parame & "  SELECT DECODE(A.MAEPRV_FECING,'',0, ROUND((((TO_DATE(TO_CHAR(sysdate,'yyyymmdd'),'YYYYMMDD') - TO_DATE((A.MAEPRV_FECING), 'YYYYMMDD'))*30)/360),0)) AS DIAS_GANADOS, "
   r_str_Parame = r_str_Parame & "  SELECT TRUNC(DECODE(A.MAEPRV_FECING,'',0, ABS((((TO_DATE(TO_CHAR(sysdate,'yyyymmdd'),'YYYYMMDD') - TO_DATE((A.MAEPRV_FECING), 'YYYYMMDD'))*30)/365))),0) AS DIAS_GANADOS, "
   
   r_str_Parame = r_str_Parame & "         (NVL(C.GESPER_DIAGOZ,0) + ROUND(NVL(D.DIAS_GOZADOS,0),0)) AS DIAS_GOZADOS, "
   r_str_Parame = r_str_Parame & "         A.MAEPRV_CODSIC, A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC, A.MAEPRV_FECING, "
   r_str_Parame = r_str_Parame & "         trunc(months_between(SYSDATE, to_date(A.MAEPRV_FECING,'YYYYMMDD'))/12) * 30 AS DIAS_VENCIDOS, "
   r_str_Parame = r_str_Parame & "         C.GESPER_CODGES, B.USUMAE_TIPJEF, "
   r_str_Parame = r_str_Parame & "         (SELECT TRIM(X.PARDES_DESCRI) FROM MNT_PARDES X WHERE X.PARDES_CODGRP = 13 AND X.PARDES_CODITE = B.USUMAE_SITUAC) AS SITUACION "
   
   r_str_Parame = r_str_Parame & "    FROM CNTBL_MAEPRV A "
   r_str_Parame = r_str_Parame & "    LEFT JOIN SEG_USUMAE B ON A.MAEPRV_CODSIC = B.USUMAE_CODSIC "
   r_str_Parame = r_str_Parame & "    LEFT JOIN CNTBL_GESPER C ON C.GESPER_TIPDOC = A.MAEPRV_TIPDOC AND C.GESPER_NUMDOC = A.MAEPRV_NUMDOC "
   r_str_Parame = r_str_Parame & "          AND C.GESPER_SITUAC = 1 AND C.GESPER_TIPTAB = 3 " '--MAESTRO
   r_str_Parame = r_str_Parame & "    LEFT JOIN (SELECT SUM(NVL(H.GESPER_IMPORT,0)) AS DIAS_GOZADOS , H.GESPER_TIPDOC, H.GESPER_NUMDOC "
   r_str_Parame = r_str_Parame & "                      FROM CNTBL_GESPER H "
   r_str_Parame = r_str_Parame & "                     Where H.GESPER_TIPTAB = 2 " '--AUTORIZADOS
   r_str_Parame = r_str_Parame & "                       AND H.GESPER_SITUAC = 2 " '--APROBADOS
   r_str_Parame = r_str_Parame & "                     GROUP BY H.GESPER_TIPDOC, H.GESPER_NUMDOC) D "
   r_str_Parame = r_str_Parame & "      ON A.MAEPRV_TIPDOC = D.GESPER_TIPDOC AND A.MAEPRV_NUMDOC = D.GESPER_NUMDOC "
   r_str_Parame = r_str_Parame & "   WHERE A.MAEPRV_TIPPER = 2  " ' --PERSONAL INTERNOS
   r_str_Parame = r_str_Parame & "     AND A.MAEPRV_SITUAC = 1 "
   
   r_str_Parame = r_str_Parame & "     AND (B.USUMAE_SITUAC = 1 OR  "
   r_str_Parame = r_str_Parame & "          (B.USUMAE_SITUAC = 2 AND (SUBSTR(B.USUMAE_FECCES,1,6) = " & r_str_PerAct & " OR SUBSTR(B.USUMAE_FECCES,1,6) = " & r_str_PerAnt & "))) "
   
   cmd_ExpVac.Enabled = True
   If fs_UserEjecutivo(modgen_g_str_CodUsu, "313") = "" And fs_UserEjecutivo(modgen_g_str_CodUsu, "314") = "" Then 'administrador vacaciones, evaluador
      'usuario comun
      cmd_ExpVac.Enabled = False
      r_str_Parame = r_str_Parame & "    AND UPPER(TRIM(B.USUMAE_CODIGO)) = '" & UCase(Trim(modgen_g_str_CodUsu)) & "'"
   ElseIf fs_UserEjecutivo(modgen_g_str_CodUsu, "314") <> "" Then
      r_str_Parame = r_str_Parame & "    AND B.USUMAE_TIPJEF = (SELECT A.USUMAE_TIPJEF FROM SEG_USUMAE A WHERE TRIM(A.USUMAE_CODIGO) = '" & UCase(Trim(modgen_g_str_CodUsu)) & "')"
   End If
   
   If (cmb_BusVac.ListIndex = 1) Then 'CODIGO PLANILLA
       If Len(Trim(txt_BusVac.Text)) > 0 Then
          r_str_Parame = r_str_Parame & "   AND TRIM(A.MAEPRV_NUMDOC) LIKE '%" & UCase(Trim(txt_BusVac.Text)) & "%'"
       End If
   ElseIf (cmb_BusVac.ListIndex = 2) Then 'NRO DOCUMENTO
       If Len(Trim(txt_BusVac.Text)) > 0 Then
          r_str_Parame = r_str_Parame & "   AND TRIM(A.MAEPRV_NUMDOC) LIKE '%" & UCase(Trim(txt_BusVac.Text)) & "%'"
       End If
   ElseIf (cmb_BusVac.ListIndex = 3) Then 'TRABAJADOR
       If Len(Trim(txt_BusVac.Text)) > 0 Then
          r_str_Parame = r_str_Parame & "   AND TRIM(A.MAEPRV_RAZSOC) LIKE '%" & UCase(Trim(txt_BusVac.Text)) & "%'"
       End If
   End If
   r_str_Parame = r_str_Parame & " ORDER BY A.MAEPRV_RAZSOC ASC "

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   grd_ListVac.Redraw = False
   r_rst_Princi.MoveFirst
   
   Do While Not r_rst_Princi.EOF
      grd_ListVac.Rows = grd_ListVac.Rows + 1
      grd_ListVac.Row = grd_ListVac.Rows - 1
      
      grd_ListVac.Col = 0
      grd_ListVac.Text = CStr(r_rst_Princi!MAEPRV_CODSIC)
      
      grd_ListVac.Col = 1
      grd_ListVac.Text = Trim(r_rst_Princi!MAEPRV_TIPDOC) & " - " & Trim(r_rst_Princi!MAEPRV_NUMDOC)
      
      grd_ListVac.Col = 3
      grd_ListVac.Text = Trim(r_rst_Princi!SITUACION & "")
      
      grd_ListVac.Col = 2
      grd_ListVac.Text = Trim(r_rst_Princi!MaePrv_RazSoc & "")
      
      If Trim(r_rst_Princi!MAEPRV_FECING & "") <> "" Then
         grd_ListVac.Col = 4
         grd_ListVac.Text = gf_FormatoFecha(r_rst_Princi!MAEPRV_FECING & "")
      End If
      
      'ganados
      grd_ListVac.Col = 5
      grd_ListVac.Text = r_rst_Princi!DIAS_GANADOS & " "
      'vencidos
      grd_ListVac.Col = 6
      grd_ListVac.Text = r_rst_Princi!DIAS_VENCIDOS & " "
      'gozados
      grd_ListVac.Col = 7
      grd_ListVac.Text = r_rst_Princi!DIAS_GOZADOS & " "
      'saldo
       grd_ListVac.Col = 8
       grd_ListVac.Text = (r_rst_Princi!DIAS_GANADOS - r_rst_Princi!DIAS_GOZADOS) & " "
      
      'vigente (saldo vencidos)
      grd_ListVac.Col = 9
      grd_ListVac.Text = r_rst_Princi!DIAS_VENCIDOS - r_rst_Princi!DIAS_GOZADOS & " "
      
      grd_ListVac.Col = 10
      grd_ListVac.Text = Trim(r_rst_Princi!MAEPRV_TIPDOC)
      
      grd_ListVac.Col = 11
      grd_ListVac.Text = Trim(r_rst_Princi!MAEPRV_NUMDOC)
      
      grd_ListVac.Col = 12
      grd_ListVac.Text = Trim(r_rst_Princi!MAEPRV_FECING & "")
      
      grd_ListVac.Col = 13
      grd_ListVac.Text = Trim(r_rst_Princi!GESPER_CODGES & "")
                  
      r_rst_Princi.MoveNext
   Loop
   
   grd_ListVac.Redraw = True
   Call gs_UbiIniGrid(grd_ListVac)
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Public Sub fs_BuscarAut()
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_str_Cadena     As String
Dim r_str_FecVct     As String
Dim r_str_FecApe     As String
Dim r_int_FecDif     As Integer
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
                    
   If fs_UserEjecutivo(modgen_g_str_CodUsu, "313") <> "" Then 'administrador de vacaciones
      cmd_AprAut.Enabled = True
      cmd_RhzAut.Enabled = True
      cmd_Reversa.Enabled = True
   Else
      If fs_UserEjecutivo(modgen_g_str_CodUsu, "314") = "" Then 'evaluador de vacaciones
         cmd_AprAut.Enabled = False
         cmd_RhzAut.Enabled = False
         cmd_Reversa.Enabled = False
         Exit Sub
      End If
   End If
   
   If cmd_AprAut.Enabled = True Then
      If cmb_BusPor.ListIndex = 1 Then
         cmd_AprAut.Enabled = False
         cmd_RhzAut.Enabled = False
         cmd_Reversa.Enabled = False
      End If
   End If

   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_ListAut)
   r_str_FecIni = Format(ipp_FecIniAut.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFinAut.Text, "yyyymmdd")
            
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT A.GESPER_CODGES, A.GESPER_FECOPE, B.MAEPRV_TIPDOC, B.MAEPRV_NUMDOC, B.MAEPRV_RAZSOC, TRIM(C.PARDES_DESCRI) AS TIPO_OPERACION, "
   r_str_Parame = r_str_Parame & "        A.GESPER_FECHA1, A.GESPER_FECHA2, A.GESPER_IMPORT, TRIM(A.GESPER_DESCRI) AS GESPER_DESCRI, A.GESPER_SITUAC, A.GESPER_TIPOPE "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_GESPER A "
   r_str_Parame = r_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.GESPER_TIPDOC AND B.MAEPRV_NUMDOC = A.GESPER_NUMDOC "
   r_str_Parame = r_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 140 AND C.PARDES_CODITE = A.GESPER_TIPOPE "
   r_str_Parame = r_str_Parame & "   LEFT JOIN SEG_USUMAE D ON D.USUMAE_CODSIC = B.MAEPRV_CODSIC "
   r_str_Parame = r_str_Parame & "  WHERE A.GESPER_TIPTAB = 2 "
   
   'r_str_Parame = r_str_Parame & "    AND A.GESPER_FECOPE BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   If cmb_BusPor.ListIndex = 0 Then
      'fecha de operaciones
      r_str_Parame = r_str_Parame & "    AND A.GESPER_FECOPE BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   ElseIf cmb_BusPor.ListIndex = 1 Then
      'fecha de goce vacaciones
      'r_str_Parame = r_str_Parame & "    AND A.GESPER_FECHA1 BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
       r_str_Parame = r_str_Parame & " AND NVL((SELECT COUNT(*) "
       r_str_Parame = r_str_Parame & "            FROM (SELECT FECHA FROM (SELECT LEVEL,TO_DATE(" & r_str_FecIni & ",'YYYYMMDD')+LEVEL-1 FECHA "
       r_str_Parame = r_str_Parame & "                    FROM DUAL "
       r_str_Parame = r_str_Parame & "                 CONNECT BY LEVEL BETWEEN 1 AND TO_NUMBER(TO_DATE(" & r_str_FecFin & ",'YYYYMMDD') - TO_DATE(" & r_str_FecIni & ",'YYYYMMDD'))+1)) "
       r_str_Parame = r_str_Parame & "           WHERE TO_NUMBER(TO_CHAR(FECHA,'YYYYMMDD')) BETWEEN A.GESPER_FECHA1 AND A.GESPER_FECHA2),0) > 0 "
   End If
   
   r_str_Parame = r_str_Parame & "  AND A.GESPER_SITUAC <> 0 "
   
   If cmb_SitAut.ListIndex <> -1 Then
      If cmb_SitAut.ItemData(cmb_SitAut.ListIndex) <> 0 Then
         r_str_Parame = r_str_Parame & "  AND A.GESPER_SITUAC =" & cmb_SitAut.ItemData(cmb_SitAut.ListIndex)
      End If
   End If
   
   If fs_UserEjecutivo(modgen_g_str_CodUsu, "313") = "" Then 'administrador de vacaciones
      r_str_Parame = r_str_Parame & "   AND D.USUMAE_TIPJEF = (SELECT X.USUMAE_TIPJEF FROM SEG_USUMAE X WHERE UPPER(TRIM(X.USUMAE_CODIGO)) = '" & UCase(Trim(modgen_g_str_CodUsu)) & "') "
   End If
   
   r_str_Parame = r_str_Parame & "  ORDER BY A.GESPER_FECOPE DESC, B.MAEPRV_RAZSOC DESC, A.GESPER_CODGES ASC "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   grd_ListAut.Redraw = False
   r_rst_Princi.MoveFirst
   
   Do While Not r_rst_Princi.EOF
      grd_ListAut.Rows = grd_ListAut.Rows + 1
      grd_ListAut.Row = grd_ListAut.Rows - 1
      
      grd_ListAut.Col = 0
      grd_ListAut.Text = CStr(r_rst_Princi!GESPER_CODGES)
      
      'grd_ListAut.Col = 1
      'grd_ListAut.Text = Trim(r_rst_Princi!MAEPRV_TIPDOC) & " - " & Trim(r_rst_Princi!maeprv_numdoc)
      
      grd_ListAut.Col = 1
      grd_ListAut.Text = gf_FormatoFecha(r_rst_Princi!GESPER_FECOPE & "")
      
      grd_ListAut.Col = 2
      grd_ListAut.Text = Trim(r_rst_Princi!MaePrv_RazSoc & "")
      
      grd_ListAut.Col = 3
      grd_ListAut.Text = Trim(r_rst_Princi!TIPO_OPERACION & "")
      
'      grd_ListAut.Col = 4
'      grd_ListAut.Text = gf_FormatoFecha(r_rst_Princi!GESPER_FECHA1)
               
      'grd_ListAut.Col = 5
      'grd_ListAut.Text = gf_FormatoFecha(r_rst_Princi!GESPER_FECHA2)
      
      If cmb_BusPor.ListIndex = 0 Then
         'BUQUEDA POR FECHA DE OPERACION
         grd_ListAut.Col = 4
         grd_ListAut.Text = gf_FormatoFecha(r_rst_Princi!GESPER_FECHA1)

         grd_ListAut.Col = 5
         grd_ListAut.Text = gf_FormatoFecha(r_rst_Princi!GESPER_FECHA2) 'FECHA HASTA
         
         grd_ListAut.Col = 6
         grd_ListAut.Text = CLng(r_rst_Princi!GESPER_IMPORT)
      Else
         'BUQUEDA POR FECHA DE GOCE
         Dim r_str_Fecha1 As String
         Dim r_str_Fecha2 As String
         
         If CLng(r_rst_Princi!GESPER_FECHA1) >= CLng(Format(ipp_FecIniAut.Text, "yyyymmdd")) Then
            grd_ListAut.Col = 4
            grd_ListAut.Text = gf_FormatoFecha(r_rst_Princi!GESPER_FECHA1) 'FECHA INICIO
            r_str_Fecha1 = grd_ListAut.Text
         Else
            grd_ListAut.Col = 4
            grd_ListAut.Text = ipp_FecIniAut.Text 'FECHA INICIO
            r_str_Fecha1 = grd_ListAut.Text
         End If
         
         If CLng(r_rst_Princi!GESPER_FECHA2) <= CLng(Format(ipp_FecFinAut.Text, "yyyymmdd")) Then
            grd_ListAut.Col = 5
            grd_ListAut.Text = gf_FormatoFecha(r_rst_Princi!GESPER_FECHA2) 'FECHA HASTA
            r_str_Fecha2 = grd_ListAut.Text
            
           'grd_ListAut.Col = 6
           'grd_ListAut.Text = CLng(r_rst_Princi!GESPER_IMPORT)
         Else
            grd_ListAut.Col = 5
            grd_ListAut.Text = ipp_FecFinAut.Text 'FECHA HASTA
            r_str_Fecha2 = grd_ListAut.Text
            
            'grd_ListAut.Col = 6
            'grd_ListAut.Text = DateDiff("d", gf_FormatoFecha(r_rst_Princi!GESPER_FECHA1), ipp_FecFinAut.Text) + 1  'DIAS SOLICITADO
         End If
         grd_ListAut.Col = 6
         grd_ListAut.Text = DateDiff("d", r_str_Fecha1, r_str_Fecha2) + 1  'DIAS SOLICITADO
      End If
                    
      'grd_ListAut.Col = 6
      'grd_ListAut.Text = CLng(r_rst_Princi!GESPER_IMPORT)
      
      grd_ListAut.Col = 7
      grd_ListAut.Text = Trim(r_rst_Princi!GESPER_DESCRI & "")
      
      grd_ListAut.Col = 8
      Select Case r_rst_Princi!GESPER_SITUAC
             Case 0: grd_ListAut.Text = "ELIMINADO"
             Case 1: grd_ListAut.Text = "PENDIENTE"
             Case 2: grd_ListAut.Text = "APROBADO"
             Case 3: grd_ListAut.Text = "RECHAZADO"
      End Select
               
      'grd_ListAut.Col = 9
      
      grd_ListAut.Col = 10
      grd_ListAut.Text = r_rst_Princi!GESPER_SITUAC
      
      grd_ListAut.Col = 11
      grd_ListAut.Text = Trim(r_rst_Princi!MAEPRV_TIPDOC)
      
      grd_ListAut.Col = 12
      grd_ListAut.Text = Trim(r_rst_Princi!MAEPRV_NUMDOC)
      
      grd_ListAut.Col = 13
      grd_ListAut.Text = Trim(r_rst_Princi!GESPER_TIPOPE)
                  
      r_rst_Princi.MoveNext
   Loop
   
   grd_ListAut.Redraw = True
   Call gs_UbiIniGrid(grd_ListAut)
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub cmb_EmpPag_Click()
   If cmb_EmpPag.ListIndex > -1 Then
      Call gs_SetFocus(cmb_SucPag)
   End If
End Sub

Private Sub cmb_EmpPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpPag_Click
   End If
End Sub

Private Sub cmb_SucPag_Click()
   Call gs_SetFocus(ipp_FecIniPag)
End Sub

Private Sub cmb_SucPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SucPag_Click
   End If
End Sub

Private Sub grd_ListAut_DblClick()
   If grd_ListAut.Rows > 0 Then
      grd_ListAut.Col = 9
      If grd_ListAut.Text = "X" Then
          grd_ListAut.Text = ""
      Else
           grd_ListAut.Text = "X"
      End If
      Call gs_RefrescaGrid(grd_ListAut)
   End If
End Sub

Private Sub grd_ListPag_DblClick()
   If grd_ListPag.Rows > 0 Then
      grd_ListPag.Col = 7
      If UCase(grd_ListPag.Text) = "NO" Then
         If grd_ListPag.TextMatrix(grd_ListPag.RowSel, 10) = "X" Then
            grd_ListPag.TextMatrix(grd_ListPag.RowSel, 10) = ""
         Else
            grd_ListPag.TextMatrix(grd_ListPag.RowSel, 10) = "X"
         End If
         grd_ListPag.TextMatrix(grd_ListPag.RowSel, 11) = ""
      Else
         If grd_ListPag.TextMatrix(grd_ListPag.RowSel, 11) = "X" Then
            grd_ListPag.TextMatrix(grd_ListPag.RowSel, 11) = ""
         Else
            grd_ListPag.TextMatrix(grd_ListPag.RowSel, 11) = "X"
         End If
         grd_ListPag.TextMatrix(grd_ListPag.RowSel, 10) = ""
      End If
      Call gs_RefrescaGrid(grd_ListPag)
   End If
End Sub

Private Sub grd_ListVac_DblClick()
   Call cmd_DetVac_Click
End Sub

Private Sub ipp_FecFinAut_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_SitAut)
   End If
End Sub

Private Sub ipp_FecFinPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusPag)
   End If
End Sub

Private Sub ipp_FecIniAut_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFinAut)
   End If
End Sub

Private Sub ipp_FecIniPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFinPag)
   End If
End Sub

Private Sub fs_GeneraAsiPag()
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_str_AsiGen        As String
Dim r_int_NumAsi        As Integer
Dim r_str_Glosa         As String
Dim r_dbl_Import        As Double
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_str_DebHab        As String
Dim r_dbl_TipSbs        As Double
Dim r_str_CadAux        As String
Dim r_int_Contad        As Integer
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
   r_str_TipNot = "D"
   r_int_NumLib = 12
   r_str_AsiGen = ""
   r_str_CadAux = ""
   r_int_Contad = 0
   
   For r_int_Contad = 0 To grd_ListPag.Rows - 1
       If grd_ListPag.TextMatrix(r_int_Contad, 7) = "NO" Then
          If grd_ListPag.TextMatrix(r_int_Contad, 10) = "X" Then
             'Inicializa variables
             r_int_NumAsi = 0
             r_str_FecPrPgoC = Format(grd_ListPag.TextMatrix(r_int_Contad, 4), "yyyymmdd")
             r_str_FecPrPgoL = grd_ListPag.TextMatrix(r_int_Contad, 4)
             
             r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, Format(grd_ListPag.TextMatrix(r_int_Contad, 4), "yyyymmdd"), 1)
             
             r_str_Glosa = ""
             r_str_Glosa = Mid((Trim(grd_ListPag.TextMatrix(r_int_Contad, 2)) & "/" & Trim(grd_ListPag.TextMatrix(r_int_Contad, 14)) & "/" & _
                           fs_ExtraeApellido(grd_ListPag.TextMatrix(r_int_Contad, 3)) & "/" & Trim(grd_ListPag.TextMatrix(r_int_Contad, 0))), 1, 60)
                                
             l_int_PerMes = modctb_int_PerMes 'Format(grd_ListPag.TextMatrix(r_int_Contad, 4), "mm")
             l_int_PerAno = modctb_int_PerAno 'Format(grd_ListPag.TextMatrix(r_int_Contad, 4), "yyyy")
             
             'Obteniendo Nro. de Asiento
             r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
             r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
                
             'Insertar en CABECERA
              Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                   r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
                  
             'Insertar en detalle
             r_dbl_MtoSol = 0
             r_dbl_MtoDol = 0
             If CInt(grd_ListPag.TextMatrix(r_int_Contad, 13)) = 1 Then 'SOLES
                r_dbl_MtoSol = CDbl(grd_ListPag.TextMatrix(r_int_Contad, 6))
                r_dbl_MtoDol = 0
             ElseIf CInt(grd_ListPag.TextMatrix(r_int_Contad, 13)) = 2 Then 'DOLARES
                r_dbl_MtoSol = Format(CDbl(CDbl(grd_ListPag.TextMatrix(r_int_Contad, 6)) * r_dbl_TipSbs), "###,###,##0.00")  'Importe * CONVERTIDO
                r_dbl_MtoDol = CDbl(grd_ListPag.TextMatrix(r_int_Contad, 6))
             End If
             
             r_str_CtaDeb = ""
             r_str_CtaHab = ""
             Select Case CInt(grd_ListPag.TextMatrix(r_int_Contad, 12))
                    Case 1: r_str_CtaDeb = "151702010101" 'Adelanto de sueldo
                    Case 2: r_str_CtaDeb = "151702010101" 'Adelanto de gratificación
                    Case 3: r_str_CtaDeb = "151702010101" 'Adelanto de vacaciones
                    Case 4: r_str_CtaDeb = "151702010101" 'Venta de vacaciones
                    Case 5: r_str_CtaDeb = "151719010101" 'Préstamo administrativo
                    Case 6: r_str_CtaDeb = "251504010101" 'Liquidaciones
                    Case 7: r_str_CtaDeb = "151702010101" 'Adelanto movilidad
                    Case 8: r_str_CtaDeb = "251419010112" 'Retencion Judicial
                    Case 9: r_str_CtaDeb = "151702010101" 'ADELANTO DE PARTICIPACION ADICIONAL A LAS UTILIDADES
             End Select
             r_str_CtaHab = "251419010109" '"111301060102"
             If CInt(grd_ListPag.TextMatrix(r_int_Contad, 13)) = 2 Then
                r_str_CtaHab = "252419010109"
             End If
                    
             Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                                  r_int_NumAsi, 1, r_str_CtaDeb, CDate(r_str_FecPrPgoL), _
                                                  r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
                                                  
             Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                                  r_int_NumAsi, 2, r_str_CtaHab, CDate(r_str_FecPrPgoL), _
                                                  r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
              
              'Actualiza flag de contabilizacion
              r_str_CadAux = ""
              r_str_CadAux = r_str_Origen & "/" & l_int_PerAno & "/" & Format(l_int_PerMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi
              g_str_Parame = ""
              g_str_Parame = g_str_Parame & " UPDATE CNTBL_GESPER  "
              g_str_Parame = g_str_Parame & "    SET GESPER_FLGCNT = 1,  "
              g_str_Parame = g_str_Parame & "        GESPER_FECCNT = " & Format(moddat_g_str_FecSis, "yyyymmdd") & ",  "
              g_str_Parame = g_str_Parame & "        GESPER_DATCNT = '" & r_str_CadAux & "'  "
              g_str_Parame = g_str_Parame & "  WHERE GESPER_CODGES  = " & CLng(grd_ListPag.TextMatrix(r_int_Contad, 0))
              
              If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                 Exit Sub
              End If
              
'             'Enviar a la tabla de autorizaciones
             g_str_Parame = ""
             g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT ( "
             g_str_Parame = g_str_Parame & " " & CLng(grd_ListPag.TextMatrix(r_int_Contad, 0)) & ", " 'COMAUT_CODOPE
             g_str_Parame = g_str_Parame & " " & Format(grd_ListPag.TextMatrix(r_int_Contad, 4), "yyyymmdd") & ", " 'COMAUT_FECOPE
             g_str_Parame = g_str_Parame & " " & grd_ListPag.TextMatrix(r_int_Contad, 15) & ", "      'COMAUT_TIPDOC
             g_str_Parame = g_str_Parame & " '" & grd_ListPag.TextMatrix(r_int_Contad, 14) & "', "    'COMAUT_NUMDOC
             g_str_Parame = g_str_Parame & " " & grd_ListPag.TextMatrix(r_int_Contad, 13) & ", "      'COMAUT_CODMON
             g_str_Parame = g_str_Parame & " " & CDbl(grd_ListPag.TextMatrix(r_int_Contad, 6)) & ", " 'COMAUT_IMPPAG
             g_str_Parame = g_str_Parame & " " & grd_ListPag.TextMatrix(r_int_Contad, 18) & ", "  'COMAUT_CODBNC
             g_str_Parame = g_str_Parame & " '" & grd_ListPag.TextMatrix(r_int_Contad, 19) & "', "  'COMAUT_CTACRR
             g_str_Parame = g_str_Parame & " '" & r_str_CtaHab & "', "  'COMAUT_CTACTB
             g_str_Parame = g_str_Parame & " '" & r_str_CadAux & "',  " 'COMAUT_DATCTB
             g_str_Parame = g_str_Parame & " '" & Trim(grd_ListPag.TextMatrix(r_int_Contad, 2)) & "',  " 'COMAUT_DESCRIPCION
             g_str_Parame = g_str_Parame & " 1,  " 'COMAUT_TIPOPE
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "  'SEGUSUCRE
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "  'SEGPLTCRE
             g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "  'SEGTERCRE
             g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "  'SEGSUCCRE

             If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                Exit Sub
             End If
              
          End If
       End If
   Next
   
   MsgBox "Se culminó proceso de generación de asientos contables para los registros." & vbCrLf & "Los asientos generados son: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
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

Private Sub fs_GenArchivo()
Dim r_int_PerAno  As Integer
Dim r_int_PerMes  As Integer
Dim r_int_NumRes  As Integer
Dim r_str_NomRes  As String
Dim r_str_Cadena  As String
Dim r_str_CadAux  As String
Dim r_dbl_PlaTot  As Double
Dim r_int_RegTot  As Integer
Dim r_int_Contad  As Integer

   r_int_PerAno = Year(moddat_g_str_FecSis)
   r_int_PerMes = Month(moddat_g_str_FecSis)
   r_str_NomRes = moddat_g_str_RutLoc & "\" & Format(moddat_g_str_FecSis, "yyyymm") & "_ADELANTOS.TXT"
                      
      'Creando Archivo
      r_int_NumRes = FreeFile
      Open r_str_NomRes For Output As r_int_NumRes
      
      r_str_Cadena = ""
      r_dbl_PlaTot = 0
      
      r_int_RegTot = 0
      For r_int_Contad = 0 To grd_ListPag.Rows - 1
          If grd_ListPag.TextMatrix(r_int_Contad, 11) = "X" Then
             r_dbl_PlaTot = r_dbl_PlaTot + CDbl(grd_ListPag.TextMatrix(r_int_Contad, 6))
             r_int_RegTot = r_int_RegTot + 1
          End If
      Next
      r_str_CadAux = ""
      For r_int_Contad = 1 To 68
          r_str_CadAux = r_str_CadAux & " "
      Next
      r_str_Cadena = r_str_Cadena & "70000110661000100040896PEN" & Format(r_dbl_PlaTot * 100, "000000000000000") & _
                                    "A" & Format(moddat_g_str_FecSis, "yyyymmdd") & "H" & "HABERES 5TA CATEGORIA    " & _
                                    Format(r_int_RegTot, "000000") & "S" & r_str_CadAux
      Print #r_int_NumRes, r_str_Cadena
      
      r_str_CadAux = ""
      For r_int_Contad = 1 To 101
          r_str_CadAux = r_str_CadAux & " "
      Next
      For r_int_Contad = 0 To grd_ListPag.Rows - 1
          If grd_ListPag.TextMatrix(r_int_Contad, 11) = "X" Then
             r_str_Cadena = ""
             r_str_Cadena = r_str_Cadena & "002" & IIf(CInt(grd_ListPag.TextMatrix(r_int_Contad, 15)) = 1, "L", "E") & _
                                           Left(Trim(grd_ListPag.TextMatrix(r_int_Contad, 14)) & "            ", 12) & _
                                           Trim(grd_ListPag.TextMatrix(r_int_Contad, 16)) & _
                                           Trim(grd_ListPag.TextMatrix(r_int_Contad, 17)) & _
                                           Left(Trim(grd_ListPag.TextMatrix(r_int_Contad, 3)) & "                                        ", 40) & _
                                           Format(grd_ListPag.TextMatrix(r_int_Contad, 6) * 100, "000000000000000") & _
                                           Left("ADELANT " & fs_nombresMes(r_int_PerMes) & "                                        ", 40) & _
                                           r_str_CadAux
             Print #r_int_NumRes, r_str_Cadena
          End If
      Next
      '                                    Left("HABERES " & fs_nombresMes(r_int_PerMes) & "                                        ", 40) & _

   'Cerrando Archivo Resumen
   Close r_int_NumRes
   
   MsgBox "Archivo Generado Exitosamente en :  " & r_str_NomRes, vbInformation, modgen_g_str_NomPlt
End Sub

Function fs_nombresMes(p_Mes As Integer) As String
   Select Case p_Mes
          Case 1: fs_nombresMes = "ENERO"
          Case 2: fs_nombresMes = "FEBRERO"
          Case 3: fs_nombresMes = "MARZO"
          Case 4: fs_nombresMes = "ABRIL"
          Case 5: fs_nombresMes = "MAYO"
          Case 6: fs_nombresMes = "JUNIO"
          Case 7: fs_nombresMes = "JULIO"
          Case 8: fs_nombresMes = "AGOSTO"
          Case 9: fs_nombresMes = "SETIEMBRE"
          Case 10: fs_nombresMes = "OCTUBRE"
          Case 11: fs_nombresMes = "NOVIEMBRE"
          Case 12: fs_nombresMes = "DICIEMBRE"
   End Select
End Function

Private Sub pln_DiaGoz_Aut_Click()
   If Len(Trim(pln_DiaGoz_Aut.Tag)) = 0 Or pln_DiaGoz_Aut.Tag = "D" Then
      pln_DiaGoz_Aut.Tag = "A"
      Call gs_SorteaGrid(grd_ListVac, 6, "N")
   Else
      pln_DiaGoz_Aut.Tag = "D"
      Call gs_SorteaGrid(grd_ListVac, 6, "N-")
   End If
End Sub

Private Sub pln_FecIng_Aut_Click()
   If Len(Trim(pln_FecIng_Aut.Tag)) = 0 Or pln_FecIng_Aut.Tag = "D" Then
      pln_FecIng_Aut.Tag = "A"
      Call gs_SorteaGrid(grd_ListVac, 12, "N")
   Else
      pln_FecIng_Aut.Tag = "D"
      Call gs_SorteaGrid(grd_ListVac, 12, "N-")
   End If
End Sub

Private Sub pln_NomTra_Aut_Click()
   If Len(Trim(pln_NomTra_Aut.Tag)) = 0 Or pln_NomTra_Aut.Tag = "D" Then
      pln_NomTra_Aut.Tag = "A"
      Call gs_SorteaGrid(grd_ListVac, 2, "C")
   Else
      pln_NomTra_Aut.Tag = "D"
      Call gs_SorteaGrid(grd_ListVac, 2, "C-")
   End If
End Sub

Private Sub pln_NroDoc_Aut_Click()
   If Len(Trim(pln_NroDoc_Aut.Tag)) = 0 Or pln_NroDoc_Aut.Tag = "D" Then
      pln_NroDoc_Aut.Tag = "A"
      Call gs_SorteaGrid(grd_ListVac, 1, "C")
   Else
      pln_NroDoc_Aut.Tag = "D"
      Call gs_SorteaGrid(grd_ListVac, 1, "C-")
   End If
End Sub

Private Sub pln_Situac_Aut_Click()
   If Len(Trim(pln_Situac_Aut.Tag)) = 0 Or pln_Situac_Aut.Tag = "D" Then
      pln_Situac_Aut.Tag = "A"
      Call gs_SorteaGrid(grd_ListVac, 3, "C")
   Else
      pln_Situac_Aut.Tag = "D"
      Call gs_SorteaGrid(grd_ListVac, 3, "C-")
   End If
End Sub

Private Sub pln_SldVen_Aut_Click()
   If Len(Trim(pln_SldVen_Aut.Tag)) = 0 Or pln_SldVen_Aut.Tag = "D" Then
      pln_SldVen_Aut.Tag = "A"
      Call gs_SorteaGrid(grd_ListVac, 9, "N")
   Else
      pln_SldVen_Aut.Tag = "D"
      Call gs_SorteaGrid(grd_ListVac, 9, "N-")
   End If
End Sub

Private Sub pln_Vencido_Aut_Click()
   If Len(Trim(pln_Vencido_Aut.Tag)) = 0 Or pln_Vencido_Aut.Tag = "D" Then
      pln_Vencido_Aut.Tag = "A"
      Call gs_SorteaGrid(grd_ListVac, 8, "N")
   Else
      pln_Vencido_Aut.Tag = "D"
      Call gs_SorteaGrid(grd_ListVac, 8, "N-")
   End If
End Sub

Private Sub pnl_CodPla_Aut_Click()
   If Len(Trim(pnl_CodPla_Aut.Tag)) = 0 Or pnl_CodPla_Aut.Tag = "D" Then
      pnl_CodPla_Aut.Tag = "A"
      Call gs_SorteaGrid(grd_ListVac, 0, "C")
   Else
      pnl_CodPla_Aut.Tag = "D"
      Call gs_SorteaGrid(grd_ListVac, 0, "C-")
   End If
End Sub

Private Sub smnu_Click(Index As Integer)
    Select Case Index
        Case 0:
              If grd_ListVac.Rows = 0 Then
                 Exit Sub
              End If
               
              If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                 Exit Sub
              End If
               
              Screen.MousePointer = 11
              Call fs_GenExc_Vac
              Screen.MousePointer = 0
        Case 1:
              If grd_ListVac.Rows = 0 Then
                 Exit Sub
              End If
               
              If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                 Exit Sub
              End If
               
              Screen.MousePointer = 11
              Call fs_GenExc_Vac_Det
              Screen.MousePointer = 0
    End Select
End Sub

 
Private Sub txt_BusPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call fs_BuscarPag
   Else
      If (cmb_BusPag.ListIndex = 1) Then
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
      Else
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
      End If
   End If
End Sub

Private Sub fs_GenExcPag()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contad        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REGISTROS DE GESTION PERSONAL"
      .Range(.Cells(2, 2), .Cells(2, 10)).Merge
      .Range(.Cells(2, 2), .Cells(2, 10)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 10)).HorizontalAlignment = xlHAlignCenter

      .Cells(4, 2) = "CÓDIGO"
      .Cells(4, 3) = "NRO DOCUMENTO"
      .Cells(4, 4) = "TIPO OPERACION"
      .Cells(4, 5) = "TRABAJADOR"
      .Cells(4, 6) = "FECHA"
      .Cells(4, 7) = "MONEDA"
      .Cells(4, 8) = "IMPORTE"
      .Cells(4, 9) = "CONTABILIZADO"
      .Cells(4, 10) = "FECHA PAGO"
         
      .Range(.Cells(4, 2), .Cells(4, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 10)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 12 'CÓDIGO
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 17 'NRO DOCUMENTO
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 30 'TIPO OPERACION"
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 40 'TRABAJADOR
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 12 'FECHA
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 21 'MONEDA
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 13 'IMPORTE
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 15 'CONTABILIZADO
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 12 'FECHA PAGO
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Size = 11
      
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_ListPag.Rows - 1
          .Cells(r_int_NumFil + 2, 2) = "'" & grd_ListPag.TextMatrix(r_int_Contad, 0) 'CÓDIGO
          .Cells(r_int_NumFil + 2, 3) = "'" & grd_ListPag.TextMatrix(r_int_Contad, 1) 'NRO DOCUMENTO
          .Cells(r_int_NumFil + 2, 4) = "'" & grd_ListPag.TextMatrix(r_int_Contad, 2) 'TIPO OPERACION"
          .Cells(r_int_NumFil + 2, 5) = "'" & grd_ListPag.TextMatrix(r_int_Contad, 3) 'TRABAJADOR
          .Cells(r_int_NumFil + 2, 6) = "'" & grd_ListPag.TextMatrix(r_int_Contad, 4) 'FECHA
          .Cells(r_int_NumFil + 2, 7) = "'" & grd_ListPag.TextMatrix(r_int_Contad, 5) 'MONEDA
          .Cells(r_int_NumFil + 2, 8) = "'" & grd_ListPag.TextMatrix(r_int_Contad, 6) 'IMPORTE
          .Cells(r_int_NumFil + 2, 9) = "'" & grd_ListPag.TextMatrix(r_int_Contad, 7) 'CONTABILIZADO
          .Cells(r_int_NumFil + 2, 10) = "'" & grd_ListPag.TextMatrix(r_int_Contad, 8) 'FECHA PAGO
                                         
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(4, 3), .Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_CalPerido_Vac_01(p_FecIng As String, p_DiaVen As Long, p_DiaGoz As Long, ByRef p_arrVac() As arr_PerVac)
Dim r_lng_FecPer As Long
Dim r_lng_ExlFil As Long
Dim r_lng_Item   As Long
Dim r_str_fecing As String
Dim r_lng_DiaGoz As Long
Dim r_dbl_Decimal As Double
Dim r_dbl_ImpAux As Double
      
      ReDim p_arrVac(0)
        
      r_lng_FecPer = 0
      If p_DiaVen > p_DiaGoz Then
         r_lng_FecPer = p_DiaVen / 30
      Else
         r_dbl_ImpAux = p_DiaGoz / 30
         r_dbl_Decimal = r_dbl_ImpAux - (CLng(r_dbl_ImpAux))
         r_lng_FecPer = (CLng(r_dbl_ImpAux))
         If r_dbl_Decimal > 0 Then
            r_lng_FecPer = r_lng_FecPer + 1
         End If
      End If
        
      r_str_fecing = p_FecIng 'fecha ingreso
      r_lng_DiaGoz = p_DiaGoz 'dias gozados
      r_lng_ExlFil = 4
      For r_lng_Item = 1 To r_lng_FecPer
          ReDim Preserve p_arrVac(UBound(p_arrVac) + 1)
          'CALCULO VENCIDOS
          p_arrVac(UBound(p_arrVac)).perVac_Item = r_lng_Item
          p_arrVac(UBound(p_arrVac)).perVac_FecIni = r_str_fecing
          r_str_fecing = DateAdd("yyyy", 1, r_str_fecing)
          p_arrVac(UBound(p_arrVac)).perVac_FecFin = r_str_fecing
          If r_lng_Item <= p_DiaVen / 30 Then
             p_arrVac(UBound(p_arrVac)).perVac_Situac = "VENCIDO"
             p_arrVac(UBound(p_arrVac)).perVac_DiaAcu = 30
          Else
             p_arrVac(UBound(p_arrVac)).perVac_Situac = "VENCIDO"
             p_arrVac(UBound(p_arrVac)).perVac_DiaAcu = 0
          End If
          'CALCULO GOZADOS
          If r_lng_DiaGoz >= p_arrVac(UBound(p_arrVac)).perVac_DiaAcu And p_arrVac(UBound(p_arrVac)).perVac_DiaAcu = 30 Then
             r_lng_DiaGoz = r_lng_DiaGoz - p_arrVac(UBound(p_arrVac)).perVac_DiaAcu
             p_arrVac(UBound(p_arrVac)).perVac_DiaGoz = 30
          Else
             p_arrVac(UBound(p_arrVac)).perVac_DiaGoz = r_lng_DiaGoz
             r_lng_DiaGoz = 0
          End If
          'CALCULO DISPONIBLE
          p_arrVac(UBound(p_arrVac)).perVac_DiaDis = p_arrVac(UBound(p_arrVac)).perVac_DiaAcu - p_arrVac(UBound(p_arrVac)).perVac_DiaGoz
      Next
End Sub

Private Sub fs_CalPerido_Vac_02(p_FecIng As String, p_DiaVen As Long, p_TipDoc As Integer, p_NumDoc As String, ByRef p_arrVac() As arr_PerVac, p_DiaSol As Long)
Dim r_lng_FecPer  As Long
Dim r_lng_ExlFil  As Long
Dim r_lng_Item    As Long
Dim r_str_fecing  As String
Dim r_lng_DiaGoz  As Long
Dim r_dbl_Decimal As Double
Dim r_dbl_ImpAux  As Double
Dim r_str_Parame  As String
Dim r_rst_Princi  As ADODB.Recordset
Dim r_dbl_TotGoz  As Double
Dim r_str_FecMax  As String
      
Dim r_arrSol()    As arr_PerVac

      ReDim p_arrVac(0)
      r_lng_FecPer = 0
      r_dbl_TotGoz = 0
      
      r_str_Parame = r_str_Parame & "SELECT SUM(A.GESPER_IMPORT) AS NUMDIA, A.GESPER_FECHA1, A.GESPER_FECHA2 "
      r_str_Parame = r_str_Parame & "  FROM CNTBL_GESPER A "
      r_str_Parame = r_str_Parame & " WHERE A.GESPER_TIPDOC = " & p_TipDoc
      r_str_Parame = r_str_Parame & "   AND TRIM(A.GESPER_NUMDOC) = '" & Trim(p_NumDoc) & "'"
      r_str_Parame = r_str_Parame & "   AND A.GESPER_TIPTAB = 4 "
      r_str_Parame = r_str_Parame & "   AND A.GESPER_SITUAC = 1 "
      r_str_Parame = r_str_Parame & " GROUP BY A.GESPER_FECHA1, A.GESPER_FECHA2 "
      r_str_Parame = r_str_Parame & " ORDER BY A.GESPER_FECHA1 ASC "
 
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
        Exit Sub
      End If
      
      r_str_FecMax = p_FecIng
      If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
         r_rst_Princi.MoveFirst
         Do While Not r_rst_Princi.EOF
            r_dbl_TotGoz = r_dbl_TotGoz + r_rst_Princi!NUMDIA
            r_str_FecMax = gf_FormatoFecha(r_rst_Princi!GESPER_FECHA2)
            r_rst_Princi.MoveNext
         Loop
      End If
      
      If p_DiaVen > r_dbl_TotGoz Then
         r_lng_FecPer = p_DiaVen / 30
      Else
         r_dbl_ImpAux = r_dbl_TotGoz / 30
         r_dbl_Decimal = r_dbl_ImpAux - (CLng(r_dbl_ImpAux))
         r_lng_FecPer = (CLng(r_dbl_ImpAux))
         If r_dbl_Decimal > 0 Then
            r_lng_FecPer = r_lng_FecPer + 1
         End If
      End If
        
      r_str_fecing = p_FecIng 'fecha ingreso
      r_lng_DiaGoz = r_dbl_TotGoz 'dias gozados
      
      'calcula periodos
      r_lng_ExlFil = 4
      For r_lng_Item = 1 To r_lng_FecPer
          ReDim Preserve p_arrVac(UBound(p_arrVac) + 1)
          p_arrVac(UBound(p_arrVac)).perVac_Item = r_lng_Item
          p_arrVac(UBound(p_arrVac)).perVac_FecIni = r_str_fecing
          r_str_fecing = DateAdd("yyyy", 1, r_str_fecing)
          p_arrVac(UBound(p_arrVac)).perVac_FecFin = r_str_fecing
          If r_lng_Item <= p_DiaVen / 30 Then
             p_arrVac(UBound(p_arrVac)).perVac_Situac = "VENCIDO"
             p_arrVac(UBound(p_arrVac)).perVac_DiaAcu = 30
          Else
             p_arrVac(UBound(p_arrVac)).perVac_Situac = "VIGENTE"
             p_arrVac(UBound(p_arrVac)).perVac_DiaAcu = 0
          End If
          p_arrVac(UBound(p_arrVac)).perVac_DiaGoz = 0
          p_arrVac(UBound(p_arrVac)).perVac_DiaDis = p_arrVac(UBound(p_arrVac)).perVac_DiaAcu - p_arrVac(UBound(p_arrVac)).perVac_DiaGoz
      Next
      
      'calcula dia gozados
      If r_dbl_TotGoz > 0 Then
         r_rst_Princi.MoveFirst
         Do While Not r_rst_Princi.EOF
            For r_lng_Item = 1 To UBound(p_arrVac)
                If Format(p_arrVac(r_lng_Item).perVac_FecIni, "yyyymmdd") = Trim(r_rst_Princi!GESPER_FECHA1) And _
                   Format(p_arrVac(r_lng_Item).perVac_FecFin, "yyyymmdd") = Trim(r_rst_Princi!GESPER_FECHA2) Then
                   p_arrVac(r_lng_Item).perVac_DiaGoz = r_rst_Princi!NUMDIA
                   p_arrVac(r_lng_Item).perVac_DiaDis = p_arrVac(r_lng_Item).perVac_DiaAcu - r_rst_Princi!NUMDIA
                   Exit For
                End If
            Next
            r_rst_Princi.MoveNext
         Loop
      End If
      
      'Ubicar Periodo disponible - dias solicitado
      Dim r_dbl_VarDis  As Long
      Dim r_lng_FilAux  As Long
      Dim r_lng_DiaAcum As Long
            
      ReDim r_arrSol(0)
      If p_DiaSol > 0 Then
         For r_lng_Item = 1 To UBound(p_arrVac)
             If p_arrVac(r_lng_Item).perVac_DiaGoz <> 30 Then
                r_dbl_VarDis = 30 - p_arrVac(r_lng_Item).perVac_DiaGoz
                
                r_lng_DiaAcum = 0
                For r_lng_FilAux = p_arrVac(r_lng_Item).perVac_DiaGoz To 29
                    r_lng_DiaAcum = r_lng_DiaAcum + 1
                    p_DiaSol = p_DiaSol - 1
                    If p_DiaSol = 0 Then
                       Exit For
                    End If
                Next
                ReDim Preserve r_arrSol(UBound(r_arrSol) + 1)
                r_arrSol(UBound(r_arrSol)).perVac_FecIni = p_arrVac(r_lng_Item).perVac_FecIni
                r_arrSol(UBound(r_arrSol)).perVac_FecFin = p_arrVac(r_lng_Item).perVac_FecFin
                r_arrSol(UBound(r_arrSol)).perVac_DiaGoz = r_lng_DiaAcum
                If p_DiaSol = 0 Then
                   Exit For
                End If
             End If
         Next
      
         'ubicar periodo acuenta
         r_lng_DiaAcum = 0
         If p_DiaSol <> 0 Then
            If r_arrSol(UBound(r_arrSol)).perVac_DiaGoz = 30 Then
               r_str_FecMax = DateAdd("yyyy", 1, r_str_FecMax)
            End If
         End If
         r_lng_FilAux = p_DiaSol
         For r_lng_Item = 1 To r_lng_FilAux
            r_lng_DiaAcum = r_lng_DiaAcum + 1
            p_DiaSol = p_DiaSol - 1
                    
            If r_lng_DiaAcum = 30 Then
               ReDim Preserve r_arrSol(UBound(r_arrSol) + 1)
               r_arrSol(UBound(r_arrSol)).perVac_FecIni = r_str_FecMax
               r_str_FecMax = DateAdd("yyyy", 1, r_str_FecMax)
               r_arrSol(UBound(r_arrSol)).perVac_FecFin = r_str_FecMax
               r_arrSol(UBound(r_arrSol)).perVac_DiaGoz = r_lng_DiaAcum
               r_lng_DiaAcum = 0
            End If
              
            If p_DiaSol = 0 Then
               If r_lng_DiaAcum <> 30 Then
                  ReDim Preserve r_arrSol(UBound(r_arrSol) + 1)
                  r_arrSol(UBound(r_arrSol)).perVac_FecIni = r_str_FecMax
                  r_str_FecMax = DateAdd("yyyy", 1, r_str_FecMax)
                  r_arrSol(UBound(r_arrSol)).perVac_FecFin = r_str_FecMax
                  r_arrSol(UBound(r_arrSol)).perVac_DiaGoz = r_lng_DiaAcum
                  r_lng_DiaAcum = 0
               End If
               Exit For
            End If
         Next
         p_arrVac = r_arrSol
      End If
End Sub

Private Sub fs_GenExc_Vac_Det()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contad        As Integer
Dim r_str_Parame        As String
Dim r_rst_Princi        As ADODB.Recordset

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "DETALLE DE VACACIONES AL " & Format(date, "DD/MM/YYYY")
            
      .Range(.Cells(2, 2), .Cells(3, 10)).Font.Bold = True
      .Range(.Cells(3, 2), .Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 3), .Cells(3, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 5), .Cells(2, 10)).Merge

      .Cells(2, 5) = "RESUMEN POR PERIODO"
      .Cells(3, 2) = "TRABAJADOR"
      .Cells(3, 3) = "NRO DOCUMENTO"
      .Cells(3, 4) = "FECHA INGRESO"
      .Cells(3, 5) = "SITUACION"
      .Cells(3, 6) = "PERIODO INICIAL"
      .Cells(3, 7) = "PERIODO FINAL"
      .Cells(3, 8) = "DIAS VENCIDOS"
      .Cells(3, 9) = "DIAS GOZADOS"
      .Cells(3, 10) = "DISPONIBLE"
         
      .Cells(2, 2).Interior.Color = RGB(146, 208, 80)
      .Cells(2, 2).Font.Bold = True
      .Range(.Cells(3, 2), .Cells(3, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 10)).Font.Bold = True

      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 50 'TRABAJADOR
      .Columns("B").HorizontalAlignment = xlHAlignLeft
      .Columns("C").ColumnWidth = 17 'NRO DOCUMENTO
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 16 'FECHA INGRESO
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 16 'SITUACION
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 16 'PERIODO DE VACACIONES
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 14 'PERIODO DE VACACIONES
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 14 'ACUMULADO
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 16 'DIAS GOZADOS
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 16 'DISPONIBLE
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(3, 10)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(3, 10)).Font.Size = 11
      
      .Cells(4, 2) = "'" & grd_ListVac.TextMatrix(grd_ListVac.Row, 2) 'TRABAJADOR
      .Cells(4, 3) = "'" & grd_ListVac.TextMatrix(grd_ListVac.Row, 1) 'NRO DOCUMENTO
      .Cells(4, 4) = "'" & grd_ListVac.TextMatrix(grd_ListVac.Row, 4) 'FECHA INGRESO 3
      Dim r_lng_FecPer As Long
      Dim r_lng_ExlFil As Long
      Dim r_lng_Item   As Long
      Dim r_str_fecing As String
      Dim r_lng_DiaGoz As Long

       ReDim l_arr_PerVac(0)
       If Len(Trim(grd_ListVac.TextMatrix(grd_ListVac.Row, 4))) > 0 Then
          Call fs_CalPerido_Vac_02(grd_ListVac.TextMatrix(grd_ListVac.Row, 4), grd_ListVac.TextMatrix(grd_ListVac.Row, 6), grd_ListVac.TextMatrix(grd_ListVac.Row, 10), _
                                   grd_ListVac.TextMatrix(grd_ListVac.Row, 11), l_arr_PerVac, 0)
       End If
       r_lng_ExlFil = 4
       For r_int_Contad = 1 To UBound(l_arr_PerVac)
           .Cells(r_lng_ExlFil, 5) = l_arr_PerVac(r_int_Contad).perVac_Situac
           .Cells(r_lng_ExlFil, 6) = "'" & l_arr_PerVac(r_int_Contad).perVac_FecIni
           .Cells(r_lng_ExlFil, 7) = "'" & l_arr_PerVac(r_int_Contad).perVac_FecFin
           .Cells(r_lng_ExlFil, 8) = l_arr_PerVac(r_int_Contad).perVac_DiaAcu
           .Cells(r_lng_ExlFil, 9) = l_arr_PerVac(r_int_Contad).perVac_DiaGoz
           .Cells(r_lng_ExlFil, 10) = l_arr_PerVac(r_int_Contad).perVac_DiaDis
       
           r_lng_ExlFil = r_lng_ExlFil + 1
       Next

      .Cells(r_lng_ExlFil, 8) = IIf(Trim(grd_ListVac.TextMatrix(grd_ListVac.Row, 6)) = "", 0, grd_ListVac.TextMatrix(grd_ListVac.Row, 6)) 'dias vencidos
      .Cells(r_lng_ExlFil, 9) = IIf(Trim(grd_ListVac.TextMatrix(grd_ListVac.Row, 7)) = "", 0, grd_ListVac.TextMatrix(grd_ListVac.Row, 7)) 'dias gozados
      .Cells(r_lng_ExlFil, 10) = .Cells(r_lng_ExlFil, 8) - .Cells(r_lng_ExlFil, 9)
      .Range(.Cells(r_lng_ExlFil, 8), .Cells(r_lng_ExlFil, 10)).Borders(xlEdgeTop).Weight = xlMedium
      r_lng_ExlFil = r_lng_ExlFil + 2
      .Range(.Cells(r_lng_ExlFil, 3), .Cells(r_lng_ExlFil, 10)).Merge
      .Cells(r_lng_ExlFil, 3) = "DETALLADO POR PERIODO"
      .Cells(r_lng_ExlFil, 3).Font.Bold = True
      
      r_lng_ExlFil = r_lng_ExlFil + 1
      .Range(.Cells(r_lng_ExlFil, 9), .Cells(r_lng_ExlFil, 10)).Merge
      .Cells(r_lng_ExlFil, 3) = "CODIGO APROB."
      .Cells(r_lng_ExlFil, 4) = "FECHA INI. GOCE."
      .Cells(r_lng_ExlFil, 5) = "FECHA FIN GOCE."
      .Cells(r_lng_ExlFil, 6) = "PERIODO INICIAL"
      .Cells(r_lng_ExlFil, 7) = "PERIODO FINAL"
      .Cells(r_lng_ExlFil, 8) = "DIAS"
      .Cells(r_lng_ExlFil, 9) = "TIPO OPERACION"

      .Range(.Cells(r_lng_ExlFil, 3), .Cells(r_lng_ExlFil, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_lng_ExlFil, 3), .Cells(r_lng_ExlFil, 10)).Font.Bold = True
      .Range(.Cells(r_lng_ExlFil, 3), .Cells(r_lng_ExlFil, 10)).HorizontalAlignment = xlHAlignCenter

      r_lng_ExlFil = r_lng_ExlFil + 1
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT A.GESPER_CODGES, A.GESPER_FECHA1 AS PERIODO_INI, A.GESPER_FECHA2 AS PERIODO_FIN, A.GESPER_IMPORT, "
      r_str_Parame = r_str_Parame & "       TRIM(B.PARDES_DESCRI) AS TIPO_OPERACION, C.GESPER_FECHA1 AS FEC_GOCE_INI, C.GESPER_FECHA2 AS FEC_GOCE_FIN "
      r_str_Parame = r_str_Parame & "  FROM CNTBL_GESPER A "
      r_str_Parame = r_str_Parame & "  LEFT JOIN MNT_PARDES B ON B.PARDES_CODGRP = 140 AND B.PARDES_CODITE = A.GESPER_TIPOPE "
      r_str_Parame = r_str_Parame & "  LEFT JOIN CNTBL_GESPER C ON C.GESPER_CODGES = A.GESPER_CODGES AND C.GESPER_TIPTAB = 2 "
      r_str_Parame = r_str_Parame & " WHERE A.GESPER_TIPTAB = 4 "
      r_str_Parame = r_str_Parame & "   AND A.GESPER_TIPDOC =  " & grd_ListVac.TextMatrix(grd_ListVac.Row, 10)
      r_str_Parame = r_str_Parame & "   AND TRIM(A.GESPER_NUMDOC) = '" & Trim(grd_ListVac.TextMatrix(grd_ListVac.Row, 11)) & "'"
      r_str_Parame = r_str_Parame & "   AND A.GESPER_SITUAC = 1 "
      r_str_Parame = r_str_Parame & " ORDER BY A.GESPER_FECHA1, A.GESPER_NUMERO ASC "

      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
        Exit Sub
      End If

      r_lng_DiaGoz = 0
      If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
         r_rst_Princi.MoveFirst
         Do While Not r_rst_Princi.EOF
            .Cells(r_lng_ExlFil, 3) = Format(r_rst_Princi!GESPER_CODGES, "0000000000")
            If Len(Trim(r_rst_Princi!FEC_GOCE_INI)) > 0 Then
               .Cells(r_lng_ExlFil, 4) = "'" & gf_FormatoFecha(r_rst_Princi!FEC_GOCE_INI)
            End If
            If Len(Trim(r_rst_Princi!FEC_GOCE_FIN)) > 0 Then
               .Cells(r_lng_ExlFil, 5) = "'" & gf_FormatoFecha(r_rst_Princi!FEC_GOCE_FIN)
            End If
            .Cells(r_lng_ExlFil, 6) = "'" & gf_FormatoFecha(r_rst_Princi!PERIODO_INI)
            .Cells(r_lng_ExlFil, 7) = "'" & gf_FormatoFecha(r_rst_Princi!PERIODO_FIN)
            .Cells(r_lng_ExlFil, 8) = r_rst_Princi!GESPER_IMPORT
            r_lng_DiaGoz = r_lng_DiaGoz + r_rst_Princi!GESPER_IMPORT
            
            .Cells(r_lng_ExlFil, 9) = Trim(r_rst_Princi!TIPO_OPERACION & "")
            .Cells(r_lng_ExlFil, 9).HorizontalAlignment = xlHAlignLeft
            r_lng_ExlFil = r_lng_ExlFil + 1
            r_rst_Princi.MoveNext
         Loop
      End If
      .Cells(r_lng_ExlFil, 8).Borders(xlEdgeTop).Weight = xlMedium
      .Cells(r_lng_ExlFil, 8) = r_lng_DiaGoz
   End With
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_Vac()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contad        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "GESTION DE VACACIONES AL " & Format(date, "DD/MM/YYYY")
      .Range(.Cells(2, 2), .Cells(2, 10)).Merge
      .Range(.Cells(2, 2), .Cells(2, 10)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 10)).HorizontalAlignment = xlHAlignCenter

      .Cells(4, 2) = "CÓDIGO PLANILLA"
      .Cells(4, 3) = "NRO DOCUMENTO"
      .Cells(4, 4) = "TRABAJADOR"
      .Cells(4, 5) = "SITUACION"
      
      .Cells(4, 6) = "FECHA DE INGRESO"
      .Cells(4, 7) = "GANADOS (DIAS)"
      .Cells(4, 8) = "GOZADOS (DIAS)"
      .Cells(4, 9) = "SALDO (DIAS)"
      .Cells(4, 10) = "VENCIDOS (DIAS)"
      .Cells(4, 11) = "SALDO VENCIDOS (DIAS)"
         
      .Range(.Cells(4, 2), .Cells(4, 11)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 11)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 17 'CÓDIGO PLANILLA
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 17 'NRO DOCUMENTO
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 50 'TRABAJADOR
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 15 'ESTADO
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 18 'FECHA INGRESO
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 16 'GANADOS (DIAS)
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 16 'GOZADOS (DIAS)
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 13 'SALDO (DIAS)
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 16 'VENCIDO (DIAS)
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 22 'SALDO VENCIDOS (DIAS)
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(11, 11)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(11, 11)).Font.Size = 11
      
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_ListVac.Rows - 1
          .Cells(r_int_NumFil + 2, 2) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 0) 'CÓDIGO PLANILLA
          .Cells(r_int_NumFil + 2, 3) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 1) 'NRO DOCUMENTO
          .Cells(r_int_NumFil + 2, 4) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 2) 'TRABAJADOR
          .Cells(r_int_NumFil + 2, 5) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 3) 'SITUACION
          .Cells(r_int_NumFil + 2, 6) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 4) 'FECHA INGRESO
          .Cells(r_int_NumFil + 2, 7) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 5) 'GANADOS (DIAS)
          .Cells(r_int_NumFil + 2, 8) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 7) 'GOZADOS (DIAS)
          .Cells(r_int_NumFil + 2, 9) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 8) 'SALDO (DIAS)
          .Cells(r_int_NumFil + 2, 10) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 6) 'VENCIDOS (DIAS)
          .Cells(r_int_NumFil + 2, 11) = "'" & grd_ListVac.TextMatrix(r_int_Contad, 9) 'VIGENTE (DIAS)
          
          r_int_NumFil = r_int_NumFil + 1
      Next
      .Range(.Cells(4, 3), .Cells(4, 11)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_Aut()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contad        As Integer
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer
Dim r_int_PerDia        As Integer
Dim r_lng_FecFin        As Long

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      
      If cmb_BusPor.ListIndex = 0 Then
         'BUQUEDA POR FECHA DE OPERACION
         .Cells(2, 2) = "GESTION DE AUTORIZACIONES DEL     " & ipp_FecIniAut.Text & "     AL     " & ipp_FecFinAut.Text & "     (BUSQUEDA FECHA OPERACION)"
      Else
         'BUQUEDA POR FECHA DE OPERACION
         .Cells(2, 2) = "GESTION DE AUTORIZACIONES DEL     " & ipp_FecIniAut.Text & "     AL     " & ipp_FecFinAut.Text & "     (BUSQUEDA FECHA GOCE VAC.)"
      End If
      
      .Range(.Cells(2, 2), .Cells(2, 10)).Merge
      .Range(.Cells(2, 2), .Cells(2, 10)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(4, 10)).HorizontalAlignment = xlHAlignCenter

      .Cells(4, 2) = "CÓDIGO INTERNO"
      .Cells(4, 3) = "FECHA OPERACION"
      .Cells(4, 4) = "TRABAJADOR"
      .Cells(4, 5) = "TIPO OPERACION"
      .Cells(4, 6) = "FECHA DESDE"
      .Cells(4, 7) = "FECHA HASTA"
      .Cells(4, 8) = "DIAS SOLICITADOS"
      .Cells(4, 9) = "COMENTARIO"
      .Cells(4, 10) = "SITUACION"
         
      .Range(.Cells(4, 2), .Cells(4, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 10)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 17 'CÓDIGO INTERNO
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 18 'FECHA OPERACION
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 47 'TRABAJADOR
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 38 'TIPO OPERACION
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 12 'FECHA DESDE
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 13 'FECHA HASTA
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 17 'DIAS SOLICITADO
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 62 'COMENTARIO
      .Columns("I").HorizontalAlignment = xlHAlignLeft
      .Columns("J").ColumnWidth = 13 'SITUACION
      .Columns("J").HorizontalAlignment = xlHAlignLeft
      
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 10)).Font.Size = 11
      
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_ListAut.Rows - 1
          .Cells(r_int_NumFil + 2, 2) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 0) 'CÓDIGO INTERNO
          .Cells(r_int_NumFil + 2, 3) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 1) 'FECHA OPERACION
          .Cells(r_int_NumFil + 2, 4) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 2) 'TRABAJADOR
          .Cells(r_int_NumFil + 2, 5) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 3) 'TIPO OPERACION
          .Cells(r_int_NumFil + 2, 6) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 4) 'FECHA DESDE
          
          'r_int_PerMes = Format(ipp_FecIniAut.Text, "mm")
          'r_int_PerAno = Format(ipp_FecIniAut.Text, "yyyy")
          'r_int_PerDia = ff_Ultimo_Dia_Mes(r_int_PerMes, r_int_PerAno)
          'r_lng_FecFin = r_int_PerAno & Format(r_int_PerMes, "00") & Format(r_int_PerDia, "00")
          
          .Cells(r_int_NumFil + 2, 7) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 5) 'FECHA HASTA
          .Cells(r_int_NumFil + 2, 8) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 6) 'DIAS SOLICITADO
          
          'If cmb_BusPor.ListIndex = 0 Then
          '   'BUQUEDA POR FECHA DE OPERACION
          '   .Cells(r_int_NumFil + 2, 7) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 5) 'FECHA HASTA
          '   .Cells(r_int_NumFil + 2, 8) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 6) 'DIAS SOLICITADO
          'Else
          '   'BUQUEDA POR FECHA DE GOCE
          '   If CLng(Format(grd_ListAut.TextMatrix(r_int_Contad, 5), "yyyymmdd")) <= r_lng_FecFin Then
          '      .Cells(r_int_NumFil + 2, 7) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 5) 'FECHA HASTA
          '      .Cells(r_int_NumFil + 2, 8) = "'" & grd_ListAut.TextMatrix(r_int_Contad, 6) 'DIAS SOLICITADO
          '   Else
          '      .Cells(r_int_NumFil + 2, 7) = "'" & gf_FormatoFecha(r_lng_FecFin) 'FECHA HASTA
          '      .Cells(r_int_NumFil + 2, 8) = "'" & DateDiff("d", grd_ListAut.TextMatrix(r_int_Contad, 4), gf_FormatoFecha(r_lng_FecFin)) + 1 'DIAS SOLICITADO
          '   End If
          'End If
          
          .Cells(r_int_NumFil + 2, 9) = "'" & Trim(grd_ListAut.TextMatrix(r_int_Contad, 7)) 'COMENTARIO
          .Cells(r_int_NumFil + 2, 10) = "'" & Trim(grd_ListAut.TextMatrix(r_int_Contad, 8)) 'SITUACION
                                         
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(4, 3), .Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub txt_BusVac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call fs_BuscarVac
   Else
      If (cmb_BusVac.ListIndex = 1) Then
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
      Else
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
      End If
   End If
End Sub
