VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Con_CreHip_10 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9540
   ClientLeft      =   2820
   ClientTop       =   1740
   ClientWidth     =   13260
   Icon            =   "GesCtb_frm_166.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13245
      _Version        =   65536
      _ExtentX        =   23363
      _ExtentY        =   16801
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
      Begin Threed.SSPanel SSPanel19 
         Height          =   6045
         Left            =   30
         TabIndex        =   17
         Top             =   3420
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
         _ExtentY        =   10663
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
         Begin TabDlg.SSTab tab_Client 
            Height          =   5925
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   13005
            _ExtentX        =   22939
            _ExtentY        =   10451
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "GesCtb_frm_166.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel2"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel11"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "GesCtb_frm_166.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel21"
            Tab(1).Control(1)=   "SSPanel29"
            Tab(1).ControlCount=   2
            TabCaption(2)   =   "Total"
            TabPicture(2)   =   "GesCtb_frm_166.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SSPanel36"
            Tab(2).Control(1)=   "SSPanel44"
            Tab(2).ControlCount=   2
            Begin Threed.SSPanel SSPanel11 
               Height          =   1065
               Left            =   30
               TabIndex        =   19
               Top             =   360
               Width           =   12915
               _Version        =   65536
               _ExtentX        =   22781
               _ExtentY        =   1879
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
               Begin Threed.SSPanel SSPanel12 
                  Height          =   285
                  Left            =   3210
                  TabIndex        =   20
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Normal"
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
               Begin MSFlexGridLib.MSFlexGrid grd_ResCal_Tit 
                  Height          =   675
                  Left            =   60
                  TabIndex        =   21
                  Top             =   360
                  Width           =   12855
                  _ExtentX        =   22675
                  _ExtentY        =   1191
                  _Version        =   393216
                  Cols            =   6
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel13 
                  Height          =   285
                  Left            =   5130
                  TabIndex        =   22
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "CPP"
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
                  Left            =   7050
                  TabIndex        =   23
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Deficiente"
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
                  Left            =   8970
                  TabIndex        =   24
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Dudoso"
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
                  Left            =   10890
                  TabIndex        =   25
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Pérdida"
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
               Begin Threed.SSPanel SSPanel20 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   26
                  Top             =   60
                  Width           =   3135
                  _Version        =   65536
                  _ExtentX        =   5530
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Total por Calificación"
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
               Height          =   4395
               Left            =   30
               TabIndex        =   27
               Top             =   1470
               Width           =   12915
               _Version        =   65536
               _ExtentX        =   22781
               _ExtentY        =   7752
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_Tit 
                  Height          =   3645
                  Left            =   60
                  TabIndex        =   28
                  Top             =   360
                  Width           =   12855
                  _ExtentX        =   22675
                  _ExtentY        =   6429
                  _Version        =   393216
                  Rows            =   20
                  Cols            =   6
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel3 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   29
                  Top             =   60
                  Width           =   4035
                  _Version        =   65536
                  _ExtentX        =   7117
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Entidad"
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
                  Left            =   4110
                  TabIndex        =   30
                  Top             =   60
                  Width           =   2505
                  _Version        =   65536
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Tipo Deuda"
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
                  Left            =   8760
                  TabIndex        =   31
                  Top             =   60
                  Width           =   2085
                  _Version        =   65536
                  _ExtentX        =   3678
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Calificación"
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
                  Left            =   6600
                  TabIndex        =   32
                  Top             =   60
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Moneda Org."
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
                  Left            =   10830
                  TabIndex        =   33
                  Top             =   60
                  Width           =   1725
                  _Version        =   65536
                  _ExtentX        =   3043
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Importe Deuda (S/.)"
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
                  Left            =   7680
                  TabIndex        =   34
                  Top             =   60
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "D. Atraso"
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
               Begin Threed.SSPanel pnl_TotDeu_Tit 
                  Height          =   285
                  Left            =   10830
                  TabIndex        =   35
                  Top             =   4020
                  Width           =   1725
                  _Version        =   65536
                  _ExtentX        =   3043
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Total Deuda ==> (S/.)"
                  Height          =   285
                  Left            =   9000
                  TabIndex        =   36
                  Top             =   4020
                  Width           =   1755
               End
            End
            Begin Threed.SSPanel SSPanel21 
               Height          =   1065
               Left            =   -74970
               TabIndex        =   37
               Top             =   360
               Width           =   12915
               _Version        =   65536
               _ExtentX        =   22781
               _ExtentY        =   1879
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
               Begin Threed.SSPanel SSPanel22 
                  Height          =   285
                  Left            =   3210
                  TabIndex        =   38
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Normal"
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
               Begin MSFlexGridLib.MSFlexGrid grd_ResCal_Cyg 
                  Height          =   675
                  Left            =   60
                  TabIndex        =   39
                  Top             =   360
                  Width           =   12855
                  _ExtentX        =   22675
                  _ExtentY        =   1191
                  _Version        =   393216
                  Cols            =   6
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel23 
                  Height          =   285
                  Left            =   5130
                  TabIndex        =   40
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "CPP"
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
               Begin Threed.SSPanel SSPanel25 
                  Height          =   285
                  Left            =   7050
                  TabIndex        =   41
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Deficiente"
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
               Begin Threed.SSPanel SSPanel26 
                  Height          =   285
                  Left            =   8970
                  TabIndex        =   42
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Dudoso"
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
                  Left            =   10890
                  TabIndex        =   43
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Pérdida"
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
                  Left            =   90
                  TabIndex        =   44
                  Top             =   60
                  Width           =   3135
                  _Version        =   65536
                  _ExtentX        =   5530
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Total por Calificación"
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
            Begin Threed.SSPanel SSPanel29 
               Height          =   4395
               Left            =   -74970
               TabIndex        =   45
               Top             =   1470
               Width           =   12915
               _Version        =   65536
               _ExtentX        =   22781
               _ExtentY        =   7752
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_Cyg 
                  Height          =   3645
                  Left            =   60
                  TabIndex        =   46
                  Top             =   360
                  Width           =   12855
                  _ExtentX        =   22675
                  _ExtentY        =   6429
                  _Version        =   393216
                  Rows            =   20
                  Cols            =   6
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel30 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   47
                  Top             =   60
                  Width           =   4035
                  _Version        =   65536
                  _ExtentX        =   7117
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Entidad"
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
               Begin Threed.SSPanel SSPanel31 
                  Height          =   285
                  Left            =   4110
                  TabIndex        =   48
                  Top             =   60
                  Width           =   2505
                  _Version        =   65536
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Tipo Deuda"
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
               Begin Threed.SSPanel SSPanel32 
                  Height          =   285
                  Left            =   8760
                  TabIndex        =   49
                  Top             =   60
                  Width           =   2085
                  _Version        =   65536
                  _ExtentX        =   3678
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Calificación"
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
               Begin Threed.SSPanel SSPanel33 
                  Height          =   285
                  Left            =   6600
                  TabIndex        =   50
                  Top             =   60
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Moneda Org."
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
                  Left            =   10830
                  TabIndex        =   51
                  Top             =   60
                  Width           =   1725
                  _Version        =   65536
                  _ExtentX        =   3043
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Importe Deuda (S/.)"
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
               Begin Threed.SSPanel SSPanel35 
                  Height          =   285
                  Left            =   7680
                  TabIndex        =   52
                  Top             =   60
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "D. Atraso"
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
               Begin Threed.SSPanel pnl_TotDeu_Cyg 
                  Height          =   285
                  Left            =   10830
                  TabIndex        =   53
                  Top             =   4020
                  Width           =   1725
                  _Version        =   65536
                  _ExtentX        =   3043
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Total Deuda ==> (S/.)"
                  Height          =   285
                  Left            =   9000
                  TabIndex        =   54
                  Top             =   4020
                  Width           =   1755
               End
            End
            Begin Threed.SSPanel SSPanel36 
               Height          =   1065
               Left            =   -74970
               TabIndex        =   58
               Top             =   360
               Width           =   12915
               _Version        =   65536
               _ExtentX        =   22781
               _ExtentY        =   1879
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
               Begin Threed.SSPanel SSPanel37 
                  Height          =   285
                  Left            =   3210
                  TabIndex        =   59
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Normal"
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
               Begin MSFlexGridLib.MSFlexGrid grd_ResCal_Tot 
                  Height          =   675
                  Left            =   60
                  TabIndex        =   60
                  Top             =   360
                  Width           =   12855
                  _ExtentX        =   22675
                  _ExtentY        =   1191
                  _Version        =   393216
                  Cols            =   6
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel38 
                  Height          =   285
                  Left            =   5130
                  TabIndex        =   61
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "CPP"
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
               Begin Threed.SSPanel SSPanel40 
                  Height          =   285
                  Left            =   7050
                  TabIndex        =   62
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Deficiente"
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
                  Left            =   8970
                  TabIndex        =   63
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Dudoso"
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
               Begin Threed.SSPanel SSPanel42 
                  Height          =   285
                  Left            =   10890
                  TabIndex        =   64
                  Top             =   60
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Pérdida"
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
               Begin Threed.SSPanel SSPanel43 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   65
                  Top             =   60
                  Width           =   3135
                  _Version        =   65536
                  _ExtentX        =   5530
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Total por Calificación"
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
            Begin Threed.SSPanel SSPanel44 
               Height          =   4395
               Left            =   -74970
               TabIndex        =   66
               Top             =   1470
               Width           =   12915
               _Version        =   65536
               _ExtentX        =   22781
               _ExtentY        =   7752
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_Tot 
                  Height          =   3645
                  Left            =   60
                  TabIndex        =   67
                  Top             =   360
                  Width           =   12855
                  _ExtentX        =   22675
                  _ExtentY        =   6429
                  _Version        =   393216
                  Rows            =   20
                  Cols            =   6
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel45 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   68
                  Top             =   60
                  Width           =   4035
                  _Version        =   65536
                  _ExtentX        =   7117
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Entidad"
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
               Begin Threed.SSPanel SSPanel46 
                  Height          =   285
                  Left            =   4110
                  TabIndex        =   69
                  Top             =   60
                  Width           =   2505
                  _Version        =   65536
                  _ExtentX        =   4419
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Tipo Deuda"
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
               Begin Threed.SSPanel SSPanel47 
                  Height          =   285
                  Left            =   8760
                  TabIndex        =   70
                  Top             =   60
                  Width           =   2085
                  _Version        =   65536
                  _ExtentX        =   3678
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Calificación"
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
               Begin Threed.SSPanel SSPanel48 
                  Height          =   285
                  Left            =   6600
                  TabIndex        =   71
                  Top             =   60
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Moneda Org."
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
               Begin Threed.SSPanel SSPanel49 
                  Height          =   285
                  Left            =   10830
                  TabIndex        =   72
                  Top             =   60
                  Width           =   1725
                  _Version        =   65536
                  _ExtentX        =   3043
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Importe Deuda (S/.)"
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
               Begin Threed.SSPanel SSPanel50 
                  Height          =   285
                  Left            =   7680
                  TabIndex        =   73
                  Top             =   60
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "D. Atraso"
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
               Begin Threed.SSPanel pnl_TotDeu_Tot 
                  Height          =   285
                  Left            =   10830
                  TabIndex        =   74
                  Top             =   4020
                  Width           =   1725
                  _Version        =   65536
                  _ExtentX        =   3043
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "0.00 "
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
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Total Deuda ==> (S/.)"
                  Height          =   285
                  Left            =   9000
                  TabIndex        =   75
                  Top             =   4020
                  Width           =   1755
               End
            End
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   750
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_166.frx":0060
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_166.frx":036A
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Buscar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12540
            Picture         =   "GesCtb_frm_166.frx":0674
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
            Height          =   315
            Left            =   720
            TabIndex        =   4
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Consulta de Crédito Hipotecario"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   315
            Left            =   720
            TabIndex        =   5
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Posición en otras Entidades Financieras"
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
            Picture         =   "GesCtb_frm_166.frx":0AB6
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1095
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
         _ExtentY        =   1931
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
            Left            =   1560
            TabIndex        =   7
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1560
            TabIndex        =   8
            Top             =   390
            Width           =   11535
            _Version        =   65536
            _ExtentX        =   20346
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_NomCyg 
            Height          =   315
            Left            =   1560
            TabIndex        =   55
            Top             =   720
            Width           =   11535
            _Version        =   65536
            _ExtentX        =   20346
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin VB.Label Label6 
            Caption         =   "Cónyuge:"
            Height          =   315
            Left            =   60
            TabIndex        =   56
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   795
         Left            =   30
         TabIndex        =   11
         Top             =   2580
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   90
            Width           =   11535
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1560
            TabIndex        =   13
            Top             =   420
            Width           =   1125
            _Version        =   196608
            _ExtentX        =   1984
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
            ButtonStyle     =   1
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
            Text            =   "0"
            MaxValue        =   "9999"
            MinValue        =   "2009"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin VB.Label Label4 
            Caption         =   "Año:"
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   420
            Width           =   1305
         End
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   285
            Left            =   60
            TabIndex        =   14
            Top             =   90
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_Con_CreHip_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_PerMes        As Integer
Dim l_int_PerAno        As Integer

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_SetFocus(cmb_PerMes)
End Sub

Private Sub grd_Listad_Cyg_Click()
   If grd_Listad_Cyg.Rows > 2 Then
      grd_Listad_Cyg.RowSel = grd_Listad_Cyg.Row
   End If
End Sub

Private Sub grd_Listad_Tit_Click()
   If grd_Listad_Tit.Rows > 2 Then
      grd_Listad_Tit.RowSel = grd_Listad_Tit.Row
   End If
End Sub

Private Sub grd_Listad_Tot_Click()
   If grd_Listad_Tot.Rows > 2 Then
      grd_Listad_Tot.RowSel = grd_Listad_Tot.Row
   End If
End Sub

Private Sub grd_ResCal_Cyg_SelChange()
   If grd_ResCal_Cyg.Rows > 2 Then
      grd_ResCal_Cyg.RowSel = grd_ResCal_Cyg.Row
   End If
End Sub

Private Sub grd_ResCal_Tit_SelChange()
   If grd_ResCal_Tit.Rows > 2 Then
      grd_ResCal_Tit.RowSel = grd_ResCal_Tit.Row
   End If
End Sub

Private Sub grd_ResCal_Tot_SelChange()
   If grd_ResCal_Tot.Rows > 2 Then
      grd_ResCal_Tot.RowSel = grd_ResCal_Tot.Row
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   pnl_NomCyg.Caption = ""
   
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   If moddat_g_int_CygTDo > 0 Then
      pnl_NomCyg.Caption = CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom
   Else
      tab_Client.TabVisible(1) = False
      tab_Client.TabVisible(2) = False
   End If
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call fs_Activa(True)

   Call fs_UltPer
   
   If l_int_PerMes > 0 Then
      Call gs_BuscarCombo_Item(cmb_PerMes, l_int_PerMes)
      
      ipp_PerAno.value = l_int_PerAno
      
      Call cmd_Buscar_Click
   End If

   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   Call fs_Buscar(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text, moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_ResCal_Tit, grd_Listad_Tit, pnl_TotDeu_Tit)
   
   If grd_ResCal_Tit.Rows = 0 Then
      MsgBox "No se encontró información registrada para el Cliente.", vbInformation, modgen_g_str_NomPlt
   End If
   
   If moddat_g_int_CygTDo > 0 Then
      Call fs_Buscar(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text, moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_ResCal_Cyg, grd_Listad_Cyg, pnl_TotDeu_Cyg)
   
      If grd_ResCal_Cyg.Rows = 0 Then
         tab_Client.TabVisible(1) = False
         tab_Client.TabVisible(2) = False
         
         MsgBox "No se encontró información registrada para el Cónyuge.", vbInformation, modgen_g_str_NomPlt
      Else
         tab_Client.TabVisible(1) = True
         tab_Client.TabVisible(2) = True
         
         Call fs_CalculaTotal
      End If
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   
   grd_ResCal_Tit.ColWidth(0) = 3125:     grd_ResCal_Tit.ColAlignment(0) = flexAlignLeftCenter
   grd_ResCal_Tit.ColWidth(1) = 1925:     grd_ResCal_Tit.ColAlignment(1) = flexAlignRightCenter
   grd_ResCal_Tit.ColWidth(2) = 1925:     grd_ResCal_Tit.ColAlignment(2) = flexAlignRightCenter
   grd_ResCal_Tit.ColWidth(3) = 1925:     grd_ResCal_Tit.ColAlignment(3) = flexAlignRightCenter
   grd_ResCal_Tit.ColWidth(4) = 1925:     grd_ResCal_Tit.ColAlignment(4) = flexAlignRightCenter
   grd_ResCal_Tit.ColWidth(5) = 1925:     grd_ResCal_Tit.ColAlignment(5) = flexAlignRightCenter
   
   grd_ResCal_Cyg.ColWidth(0) = 3125:     grd_ResCal_Cyg.ColAlignment(0) = flexAlignLeftCenter
   grd_ResCal_Cyg.ColWidth(1) = 1925:     grd_ResCal_Cyg.ColAlignment(1) = flexAlignRightCenter
   grd_ResCal_Cyg.ColWidth(2) = 1925:     grd_ResCal_Cyg.ColAlignment(2) = flexAlignRightCenter
   grd_ResCal_Cyg.ColWidth(3) = 1925:     grd_ResCal_Cyg.ColAlignment(3) = flexAlignRightCenter
   grd_ResCal_Cyg.ColWidth(4) = 1925:     grd_ResCal_Cyg.ColAlignment(4) = flexAlignRightCenter
   grd_ResCal_Cyg.ColWidth(5) = 1925:     grd_ResCal_Cyg.ColAlignment(5) = flexAlignRightCenter
   
   grd_ResCal_Tot.ColWidth(0) = 3125:     grd_ResCal_Tot.ColAlignment(0) = flexAlignLeftCenter
   grd_ResCal_Tot.ColWidth(1) = 1925:     grd_ResCal_Tot.ColAlignment(1) = flexAlignRightCenter
   grd_ResCal_Tot.ColWidth(2) = 1925:     grd_ResCal_Tot.ColAlignment(2) = flexAlignRightCenter
   grd_ResCal_Tot.ColWidth(3) = 1925:     grd_ResCal_Tot.ColAlignment(3) = flexAlignRightCenter
   grd_ResCal_Tot.ColWidth(4) = 1925:     grd_ResCal_Tot.ColAlignment(4) = flexAlignRightCenter
   grd_ResCal_Tot.ColWidth(5) = 1925:     grd_ResCal_Tot.ColAlignment(5) = flexAlignRightCenter
   
   grd_Listad_Tit.ColWidth(0) = 4025:     grd_Listad_Tit.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_Tit.ColWidth(1) = 2495:     grd_Listad_Tit.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad_Tit.ColWidth(2) = 1085:     grd_Listad_Tit.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad_Tit.ColWidth(3) = 1085:     grd_Listad_Tit.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_Tit.ColWidth(4) = 2075:     grd_Listad_Tit.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad_Tit.ColWidth(5) = 1715:     grd_Listad_Tit.ColAlignment(5) = flexAlignRightCenter

   grd_Listad_Cyg.ColWidth(0) = 4025:     grd_Listad_Cyg.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_Cyg.ColWidth(1) = 2495:     grd_Listad_Cyg.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad_Cyg.ColWidth(2) = 1085:     grd_Listad_Cyg.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad_Cyg.ColWidth(3) = 1085:     grd_Listad_Cyg.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_Cyg.ColWidth(4) = 2075:     grd_Listad_Cyg.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad_Cyg.ColWidth(5) = 1715:     grd_Listad_Cyg.ColAlignment(5) = flexAlignRightCenter

   grd_Listad_Tot.ColWidth(0) = 4025:     grd_Listad_Tot.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_Tot.ColWidth(1) = 2495:     grd_Listad_Tot.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad_Tot.ColWidth(2) = 1085:     grd_Listad_Tot.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad_Tot.ColWidth(3) = 1085:     grd_Listad_Tot.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_Tot.ColWidth(4) = 2075:     grd_Listad_Tot.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad_Tot.ColWidth(5) = 1715:     grd_Listad_Tot.ColAlignment(5) = flexAlignRightCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_ResCal_Tit)
   Call gs_LimpiaGrid(grd_ResCal_Cyg)
   Call gs_LimpiaGrid(grd_ResCal_Tot)
   
   Call gs_LimpiaGrid(grd_Listad_Tit)
   Call gs_LimpiaGrid(grd_Listad_Cyg)
   Call gs_LimpiaGrid(grd_Listad_Tot)
   
   pnl_TotDeu_Tit.Caption = "0.00 "
   pnl_TotDeu_Cyg.Caption = "0.00 "
   pnl_TotDeu_Tot.Caption = "0.00 "
   
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = Format(Year(Date), "0000")
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_PerMes.Enabled = p_Activa
   ipp_PerAno.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa

   grd_ResCal_Tit.Enabled = Not p_Activa
   grd_ResCal_Cyg.Enabled = Not p_Activa
   grd_ResCal_Tot.Enabled = Not p_Activa
   
   grd_Listad_Tit.Enabled = Not p_Activa
   grd_Listad_Cyg.Enabled = Not p_Activa
   grd_Listad_Tot.Enabled = Not p_Activa
End Sub

Private Sub fs_UltPer()
   l_int_PerMes = 0
   l_int_PerAno = 0

   'Obteniendo Datos de Reusmen
   g_str_Parame = "SELECT * FROM CLI_RCCCAB WHERE "
   g_str_Parame = g_str_Parame & "RCCCAB_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "RCCCAB_NUMDOC = '" & moddat_g_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "ORDER BY RCCCAB_PERANO DESC, RCCCAB_PERMES DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   l_int_PerMes = g_rst_Princi("RCCCAB_PERMES")
   l_int_PerAno = g_rst_Princi("RCCCAB_PERANO")

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, p_ResCal As MSFlexGrid, ByVal p_Listad As MSFlexGrid, ByVal p_TotDeu As SSPanel)
   Dim r_dbl_TotDeu        As Double
   Dim r_int_Contad        As Integer
   Dim r_str_CodEmp        As String
   Dim r_str_NomEmp        As String
   Dim r_int_FlgNom        As Integer
   Dim r_dbl_DeuPar        As Double
   
   Call gs_LimpiaGrid(p_ResCal)
   Call gs_LimpiaGrid(p_Listad)
   
   p_TotDeu.Caption = "0.00 "
   
   'Obteniendo Datos de Reusmen
   g_str_Parame = "SELECT * FROM CLI_RCCCAB WHERE "
   g_str_Parame = g_str_Parame & "RCCCAB_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "RCCCAB_NUMDOC = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "RCCCAB_PERMES = " & CStr(p_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "RCCCAB_PERANO = " & CStr(p_PerAno) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   p_ResCal.Redraw = False
      
   g_rst_Princi.MoveFirst
   
   p_ResCal.Rows = p_ResCal.Rows + 1
   p_ResCal.Row = p_ResCal.Rows - 1
   
   p_ResCal.Col = 0:          p_ResCal.Text = "EN DINERO (S/.)"
   p_ResCal.Col = 1:          p_ResCal.Text = Format(g_rst_Princi("RCCCAB_DEUCA0"), "###,###,##0.00")
   p_ResCal.Col = 2:          p_ResCal.Text = Format(g_rst_Princi("RCCCAB_DEUCA1"), "###,###,##0.00")
   p_ResCal.Col = 3:          p_ResCal.Text = Format(g_rst_Princi("RCCCAB_DEUCA2"), "###,###,##0.00")
   p_ResCal.Col = 4:          p_ResCal.Text = Format(g_rst_Princi("RCCCAB_DEUCA3"), "###,###,##0.00")
   p_ResCal.Col = 5:          p_ResCal.Text = Format(g_rst_Princi("RCCCAB_DEUCA4"), "###,###,##0.00")
      
      
   r_dbl_TotDeu = 0
   
   For r_int_Contad = 0 To 4
      r_dbl_TotDeu = r_dbl_TotDeu + g_rst_Princi("RCCCAB_DEUCA" & CStr(r_int_Contad))
   Next r_int_Contad
   
   p_TotDeu.Caption = Format(r_dbl_TotDeu, "###,###,##0.00") & " "
   
   If r_dbl_TotDeu = 0 Then
      r_dbl_TotDeu = 1
   End If
      
   p_ResCal.Rows = p_ResCal.Rows + 1
   p_ResCal.Row = p_ResCal.Rows - 1
   
   p_ResCal.Col = 0:          p_ResCal.Text = "EN PORCENTAJE (%)"
   p_ResCal.Col = 1:          p_ResCal.Text = Format(g_rst_Princi("RCCCAB_DEUCA0") / r_dbl_TotDeu * 100, "##0.00")
   p_ResCal.Col = 2:          p_ResCal.Text = Format(g_rst_Princi("RCCCAB_DEUCA1") / r_dbl_TotDeu * 100, "##0.00")
   p_ResCal.Col = 3:          p_ResCal.Text = Format(g_rst_Princi("RCCCAB_DEUCA2") / r_dbl_TotDeu * 100, "##0.00")
   p_ResCal.Col = 4:          p_ResCal.Text = Format(g_rst_Princi("RCCCAB_DEUCA3") / r_dbl_TotDeu * 100, "##0.00")
   p_ResCal.Col = 5:          p_ResCal.Text = Format(g_rst_Princi("RCCCAB_DEUCA4") / r_dbl_TotDeu * 100, "##0.00")
      
   p_ResCal.Redraw = True
   Call gs_UbiIniGrid(p_ResCal)
   
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Obteniendo Detalle
   g_str_Parame = "SELECT * FROM CLI_RCCDET WHERE "
   g_str_Parame = g_str_Parame & "RCCDET_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "RCCDET_NUMDOC = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "RCCDET_PERMES = " & CStr(p_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "RCCDET_PERANO = " & CStr(p_PerAno) & " "
   g_str_Parame = g_str_Parame & "ORDER BY RCCDET_CODEMP ASC, RCCDET_CLASIF DESC, RCCDET_MTOSOL+RCCDET_MTODOL DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      p_Listad.Redraw = False
   
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodEmp = CStr(g_rst_Princi!RCCDET_CODEMP)
         r_str_NomEmp = moddat_gf_Consulta_NomEntFin(g_rst_Princi!RCCDET_CODEMP)
         
         r_int_FlgNom = 1
         r_dbl_DeuPar = 0
         
         Do While Not g_rst_Princi.EOF And r_str_CodEmp = CStr(g_rst_Princi!RCCDET_CODEMP)
            p_Listad.Rows = p_Listad.Rows + 1
            p_Listad.Row = p_Listad.Rows - 1
            
            If r_int_FlgNom = 1 Then
               p_Listad.Col = 0
               p_Listad.Text = IIf(Len(Trim(r_str_NomEmp)) > 0, r_str_NomEmp, CStr(g_rst_Princi!RCCDET_CODEMP))
               
               r_int_FlgNom = 2
            End If
         
            p_Listad.Col = 1
            p_Listad.Text = moddat_gf_Consulta_ParDes("264", CStr(g_rst_Princi!RCCDET_TIPDEU))
            
            p_Listad.Col = 2
            p_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!RCCDET_MONDEU))
            
            p_Listad.Col = 3
            p_Listad.Text = CStr(g_rst_Princi!RCCDET_DIAATR)
            
            p_Listad.Col = 4
            p_Listad.Text = moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!RCCDET_CLASIF + 1))
         
            p_Listad.Col = 5
            p_Listad.Text = IIf(g_rst_Princi!RCCDET_MTOSOL > 0, Format(g_rst_Princi!RCCDET_MTOSOL, "###,###,##0.00"), Format(g_rst_Princi!RCCDET_MTODOL, "###,###,##0.00"))
            
            r_dbl_DeuPar = r_dbl_DeuPar + CDbl(p_Listad.Text)
         
            g_rst_Princi.MoveNext
         
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
         
         p_Listad.Rows = p_Listad.Rows + 1
         p_Listad.Row = p_Listad.Rows - 1
         
         p_Listad.Col = 4
         p_Listad.CellAlignment = flexAlignRightCenter
         p_Listad.CellForeColor = modgen_g_con_ColRoj
         p_Listad.Text = "TOTAL DEUDA ==>"
         
         p_Listad.Col = 5
         p_Listad.CellForeColor = modgen_g_con_ColRoj
         p_Listad.Text = Format(r_dbl_DeuPar, "###,###,##0.00")
         
         p_Listad.Rows = p_Listad.Rows + 1
         
         If g_rst_Princi.EOF Then
            Exit Do
         End If
      Loop
   
      p_Listad.Redraw = True
      Call gs_UbiIniGrid(p_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_CalculaTotal()
   Dim r_dbl_TotCa0     As Double
   Dim r_dbl_TotCa1     As Double
   Dim r_dbl_TotCa2     As Double
   Dim r_dbl_TotCa3     As Double
   Dim r_dbl_TotCa4     As Double
   Dim r_dbl_TotDeu     As Double
   Dim r_str_CodEmp     As String
   Dim r_str_NomEmp     As String
   Dim r_int_FlgNom     As Integer
   Dim r_dbl_DeuPar     As Double
   Dim r_str_TipCli     As String

   Call gs_LimpiaGrid(grd_ResCal_Tot)
   pnl_TotDeu_Tot.Caption = "0.00 "
   
   grd_ResCal_Tit.Redraw = False
   grd_ResCal_Cyg.Redraw = False
   
   grd_ResCal_Tit.Row = 0
   grd_ResCal_Cyg.Row = 0
   
   grd_ResCal_Tit.Col = 1:    r_dbl_TotCa0 = r_dbl_TotCa0 + CDbl(grd_ResCal_Tit.Text)
   grd_ResCal_Cyg.Col = 1:    r_dbl_TotCa0 = r_dbl_TotCa0 + CDbl(grd_ResCal_Cyg.Text)
   
   grd_ResCal_Tit.Col = 2:    r_dbl_TotCa1 = r_dbl_TotCa1 + CDbl(grd_ResCal_Tit.Text)
   grd_ResCal_Cyg.Col = 2:    r_dbl_TotCa1 = r_dbl_TotCa1 + CDbl(grd_ResCal_Cyg.Text)
   
   grd_ResCal_Tit.Col = 3:    r_dbl_TotCa2 = r_dbl_TotCa2 + CDbl(grd_ResCal_Tit.Text)
   grd_ResCal_Cyg.Col = 3:    r_dbl_TotCa2 = r_dbl_TotCa2 + CDbl(grd_ResCal_Cyg.Text)
   
   grd_ResCal_Tit.Col = 4:    r_dbl_TotCa3 = r_dbl_TotCa3 + CDbl(grd_ResCal_Tit.Text)
   grd_ResCal_Cyg.Col = 4:    r_dbl_TotCa3 = r_dbl_TotCa3 + CDbl(grd_ResCal_Cyg.Text)
   
   grd_ResCal_Tit.Col = 5:    r_dbl_TotCa4 = r_dbl_TotCa4 + CDbl(grd_ResCal_Tit.Text)
   grd_ResCal_Cyg.Col = 5:    r_dbl_TotCa4 = r_dbl_TotCa4 + CDbl(grd_ResCal_Cyg.Text)
   
   r_dbl_TotDeu = r_dbl_TotCa0 + r_dbl_TotCa1 + r_dbl_TotCa2 + r_dbl_TotCa3 + r_dbl_TotCa4
   
   pnl_TotDeu_Tot.Caption = Format(r_dbl_TotDeu, "###,###,##0.00") & " "
   
   grd_ResCal_Tit.Redraw = True
   grd_ResCal_Cyg.Redraw = True
   
   Call gs_UbiIniGrid(grd_ResCal_Tit)
   Call gs_UbiIniGrid(grd_ResCal_Cyg)
   
   'Resumen Total
   grd_ResCal_Tot.Redraw = False
   
   grd_ResCal_Tot.Rows = grd_ResCal_Tot.Rows + 1
   grd_ResCal_Tot.Row = grd_ResCal_Tot.Rows - 1
   
   grd_ResCal_Tot.Col = 0:          grd_ResCal_Tot.Text = "EN DINERO (S/.)"
   grd_ResCal_Tot.Col = 1:          grd_ResCal_Tot.Text = Format(r_dbl_TotCa0, "###,###,##0.00")
   grd_ResCal_Tot.Col = 2:          grd_ResCal_Tot.Text = Format(r_dbl_TotCa1, "###,###,##0.00")
   grd_ResCal_Tot.Col = 3:          grd_ResCal_Tot.Text = Format(r_dbl_TotCa2, "###,###,##0.00")
   grd_ResCal_Tot.Col = 4:          grd_ResCal_Tot.Text = Format(r_dbl_TotCa3, "###,###,##0.00")
   grd_ResCal_Tot.Col = 5:          grd_ResCal_Tot.Text = Format(r_dbl_TotCa4, "###,###,##0.00")
   
   
   grd_ResCal_Tot.Rows = grd_ResCal_Tot.Rows + 1
   grd_ResCal_Tot.Row = grd_ResCal_Tot.Rows - 1
   
   grd_ResCal_Tot.Col = 0:          grd_ResCal_Tot.Text = "EN PORCENTAJE (%)"
   grd_ResCal_Tot.Col = 1:          grd_ResCal_Tot.Text = Format(r_dbl_TotCa0 / r_dbl_TotDeu * 100, "##0.00")
   grd_ResCal_Tot.Col = 2:          grd_ResCal_Tot.Text = Format(r_dbl_TotCa1 / r_dbl_TotDeu * 100, "##0.00")
   grd_ResCal_Tot.Col = 3:          grd_ResCal_Tot.Text = Format(r_dbl_TotCa2 / r_dbl_TotDeu * 100, "##0.00")
   grd_ResCal_Tot.Col = 4:          grd_ResCal_Tot.Text = Format(r_dbl_TotCa3 / r_dbl_TotDeu * 100, "##0.00")
   grd_ResCal_Tot.Col = 5:          grd_ResCal_Tot.Text = Format(r_dbl_TotCa4 / r_dbl_TotDeu * 100, "##0.00")
   
   grd_ResCal_Tot.Redraw = True
   
   Call gs_UbiIniGrid(grd_ResCal_Tot)
   

   'Obteniendo Detalle
   g_str_Parame = "SELECT * FROM CLI_RCCDET WHERE "
   g_str_Parame = g_str_Parame & "((RCCDET_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND RCCDET_NUMDOC = '" & moddat_g_str_NumDoc & "') OR "
   g_str_Parame = g_str_Parame & " (RCCDET_TIPDOC = " & CStr(moddat_g_int_CygTDo) & " AND RCCDET_NUMDOC = '" & moddat_g_str_CygNDo & "')) AND "
   g_str_Parame = g_str_Parame & "RCCDET_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "RCCDET_PERANO = " & CStr(ipp_PerAno.value) & " "
   g_str_Parame = g_str_Parame & "ORDER BY RCCDET_CODEMP ASC, RCCDET_CLASIF DESC, RCCDET_MTOSOL+RCCDET_MTODOL DESC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad_Tot.Redraw = False
   
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_CodEmp = CStr(g_rst_Princi!RCCDET_CODEMP)
         r_str_NomEmp = moddat_gf_Consulta_NomEntFin(g_rst_Princi!RCCDET_CODEMP)
         
         r_int_FlgNom = 1
         r_dbl_DeuPar = 0
         
         Do While Not g_rst_Princi.EOF And r_str_CodEmp = CStr(g_rst_Princi!RCCDET_CODEMP)
            grd_Listad_Tot.Rows = grd_Listad_Tot.Rows + 1
            grd_Listad_Tot.Row = grd_Listad_Tot.Rows - 1
            
            If r_int_FlgNom = 1 Then
               grd_Listad_Tot.Col = 0
               grd_Listad_Tot.Text = IIf(Len(Trim(r_str_NomEmp)) > 0, r_str_NomEmp, CStr(g_rst_Princi!RCCDET_CODEMP))
               
               r_int_FlgNom = 2
            End If
            
            If g_rst_Princi!RCCDET_TIPDOC = moddat_g_int_TipDoc And Trim(g_rst_Princi!RCCDET_NUMDOC) = moddat_g_str_NumDoc Then
               r_str_TipCli = "(T) - "
            Else
               r_str_TipCli = "(C) - "
            End If
            
            grd_Listad_Tot.Col = 1
            grd_Listad_Tot.Text = r_str_TipCli & moddat_gf_Consulta_ParDes("264", CStr(g_rst_Princi!RCCDET_TIPDEU))
            
            grd_Listad_Tot.Col = 2
            grd_Listad_Tot.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!RCCDET_MONDEU))
            
            grd_Listad_Tot.Col = 3
            grd_Listad_Tot.Text = CStr(g_rst_Princi!RCCDET_DIAATR)
            
            grd_Listad_Tot.Col = 4
            grd_Listad_Tot.Text = moddat_gf_Consulta_ParDes("058", CStr(g_rst_Princi!RCCDET_CLASIF + 1))
         
            grd_Listad_Tot.Col = 5
            grd_Listad_Tot.Text = IIf(g_rst_Princi!RCCDET_MTOSOL > 0, Format(g_rst_Princi!RCCDET_MTOSOL, "###,###,##0.00"), Format(g_rst_Princi!RCCDET_MTODOL, "###,###,##0.00"))
            
            r_dbl_DeuPar = r_dbl_DeuPar + CDbl(grd_Listad_Tot.Text)
         
            g_rst_Princi.MoveNext
         
            If g_rst_Princi.EOF Then
               Exit Do
            End If
         Loop
         
         grd_Listad_Tot.Rows = grd_Listad_Tot.Rows + 1
         grd_Listad_Tot.Row = grd_Listad_Tot.Rows - 1
         
         grd_Listad_Tot.Col = 4
         grd_Listad_Tot.CellAlignment = flexAlignRightCenter
         grd_Listad_Tot.CellForeColor = modgen_g_con_ColRoj
         grd_Listad_Tot.Text = "TOTAL DEUDA ==>"
         
         grd_Listad_Tot.Col = 5
         grd_Listad_Tot.CellForeColor = modgen_g_con_ColRoj
         grd_Listad_Tot.Text = Format(r_dbl_DeuPar, "###,###,##0.00")
         
         grd_Listad_Tot.Rows = grd_Listad_Tot.Rows + 1
         
         If g_rst_Princi.EOF Then
            Exit Do
         End If
      Loop
   
      grd_Listad_Tot.Redraw = True
      Call gs_UbiIniGrid(grd_Listad_Tot)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
