VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_EntRen_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17265
   Icon            =   "GesCtb_frm_192.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   17265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8745
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   17265
      _Version        =   65536
      _ExtentX        =   30454
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   18
         Top             =   60
         Width           =   17145
         _Version        =   65536
         _ExtentX        =   30242
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
            TabIndex        =   19
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Registros Entregas a Rendir"
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
            Picture         =   "GesCtb_frm_192.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   20
         Top             =   760
         Width           =   17145
         _Version        =   65536
         _ExtentX        =   30242
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
         Begin VB.CommandButton cmd_Procesar 
            Height          =   585
            Left            =   4830
            Picture         =   "GesCtb_frm_192.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Procesar Registros"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1200
            Picture         =   "GesCtb_frm_192.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_192.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_192.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   16530
            Picture         =   "GesCtb_frm_192.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_192.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   600
            Picture         =   "GesCtb_frm_192.frx":168A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   3030
            Picture         =   "GesCtb_frm_192.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   4230
            Picture         =   "GesCtb_frm_192.frx":1C9E
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Detalle 
            Height          =   585
            Left            =   3630
            Picture         =   "GesCtb_frm_192.frx":1FA8
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Detalle"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   825
         Left            =   60
         TabIndex        =   21
         Top             =   1440
         Width           =   17145
         _Version        =   65536
         _ExtentX        =   30242
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
         Begin VB.CheckBox chk_Estado 
            Caption         =   "Incluir los Procesados"
            Height          =   195
            Left            =   10470
            TabIndex        =   4
            Top             =   480
            Width           =   2925
         End
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   3465
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   3465
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   6780
            TabIndex        =   2
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   8160
            TabIndex        =   3
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Caja"
            Height          =   195
            Left            =   5520
            TabIndex        =   24
            Top             =   450
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sucursal:"
            Height          =   195
            Left            =   180
            TabIndex        =   23
            Top             =   450
            Width           =   660
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   22
            Top             =   120
            Width           =   660
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   630
         Left            =   60
         TabIndex        =   25
         Top             =   7880
         Width           =   17145
         _Version        =   65536
         _ExtentX        =   30242
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
         Begin VB.TextBox txt_Buscar 
            Height          =   315
            Left            =   5400
            MaxLength       =   100
            TabIndex        =   16
            Top             =   180
            Width           =   4425
         End
         Begin VB.ComboBox cmb_Buscar 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   180
            Width           =   2595
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Columna a Buscar:"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   27
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Por:"
            Height          =   195
            Left            =   4530
            TabIndex        =   26
            Top             =   240
            Width           =   825
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   5535
         Left            =   60
         TabIndex        =   28
         Top             =   2295
         Width           =   17145
         _Version        =   65536
         _ExtentX        =   30242
         _ExtentY        =   9763
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
         Begin Threed.SSPanel pnl_Proces 
            Height          =   285
            Left            =   12870
            TabIndex        =   34
            Top             =   60
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1094
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Proces."
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
         Begin Threed.SSPanel pnl_DiaPte 
            Height          =   285
            Left            =   13470
            TabIndex        =   37
            Top             =   60
            Width           =   705
            _Version        =   65536
            _ExtentX        =   1235
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Dias Pte."
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
            Left            =   60
            TabIndex        =   40
            Top             =   60
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
         Begin Threed.SSPanel pnl_Respon 
            Height          =   285
            Left            =   3465
            TabIndex        =   39
            Top             =   60
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5115
            Left            =   30
            TabIndex        =   29
            Top             =   360
            Width           =   17090
            _ExtentX        =   30136
            _ExtentY        =   9022
            _Version        =   393216
            Rows            =   24
            Cols            =   25
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   285
            Left            =   9105
            TabIndex        =   30
            Top             =   60
            Width           =   675
            _Version        =   65536
            _ExtentX        =   1191
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
            Left            =   9765
            TabIndex        =   31
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1940
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
         Begin Threed.SSPanel pnl_Selecc 
            Height          =   285
            Left            =   16020
            TabIndex        =   32
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
               Left            =   820
               TabIndex        =   33
               Top             =   10
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_MtoRen 
            Height          =   285
            Left            =   10845
            TabIndex        =   35
            Top             =   60
            Width           =   1020
            _Version        =   65536
            _ExtentX        =   1799
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto Rendido"
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
            Left            =   1200
            TabIndex        =   38
            Top             =   60
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
            Left            =   7425
            TabIndex        =   41
            Top             =   60
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
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
            Left            =   5490
            TabIndex        =   42
            Top             =   60
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Beneficiario"
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
            Left            =   14145
            TabIndex        =   43
            Top             =   60
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1834
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
            Left            =   15150
            TabIndex        =   44
            Top             =   60
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1552
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cód.Pago"
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
            Left            =   2220
            TabIndex        =   45
            Top             =   60
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
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
         Begin Threed.SSPanel pnl_Saldo 
            Height          =   285
            Left            =   11835
            TabIndex        =   36
            Top             =   60
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1852
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto Saldo"
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
Attribute VB_Name = "frm_Ctb_EntRen_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_arr_Sucurs()      As moddat_tpo_Genera

Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 10)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 14) = ""
             End If
         Next r_Fila
      End If
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             If UCase(grd_Listad.TextMatrix(r_Fila, 10)) = "NO" Then
                grd_Listad.TextMatrix(r_Fila, 14) = "X"
             End If
         Next r_Fila
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub cmb_Buscar_Click()
    If (cmb_Buscar.ListIndex = 0 Or cmb_Buscar.ListIndex = -1) Then
        txt_Buscar.Enabled = False
        Call gs_SetFocus(cmd_Buscar)
    Else
        txt_Buscar.Enabled = True
        Call gs_SetFocus(txt_Buscar)
    End If
    txt_Buscar.Text = ""
End Sub

Private Sub cmb_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_Buscar.Enabled = False) Then
          Call gs_SetFocus(cmd_Buscar)
      Else
          Call gs_SetFocus(txt_Buscar)
      End If
   End If
End Sub

Private Sub cmb_Empres_Click()
   If cmb_Empres.ListIndex > -1 Then
      Screen.MousePointer = 11
      
      moddat_g_str_CodEmp = l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo
      moddat_g_str_RazSoc = cmb_Empres.Text
      
      Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   
      cmb_Sucurs.ListIndex = 0
      Call gs_SetFocus(cmb_Sucurs)
      Screen.MousePointer = 0
   Else
      cmb_Sucurs.Clear
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmb_Sucurs_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Sucurs_Click
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1 'insert
   frm_Ctb_EntRen_02.Show 1
End Sub

Private Sub cmd_Borrar_Click()
Dim r_str_AsiGen   As String
Dim r_dbl_ImpAux   As Double

   r_str_AsiGen = ""
   moddat_g_str_Codigo = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If UCase(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 10))) = "SI" Then
      MsgBox "El registro esta procesado, no se pudo eliminar.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 21)) <> "" Then
     MsgBox "El registro esta asociado, no se pudo eliminar.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 22)) <> "" Then
     MsgBox "El registro tiene asociado otros registros, no se pudo eliminar.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
    
   moddat_g_str_Codigo = CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   '---------------------------------
   'procesado por Compensasion
   If fs_ValMod_Aut(moddat_g_str_Codigo, 1) = False Then
      Exit Sub
   End If
       
   'validar importe rendido - nro filas detalle
   If CDbl(CStr(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 8)))) <> CDbl(CStr(0)) Or CLng(grd_Listad.TextMatrix(grd_Listad.Row, 23)) > 0 Then
      MsgBox "Tiene que eliminar su detalle, no se puede eliminar.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   If (Format(moddat_g_str_FecSis, "yyyymmdd") < Format(modctb_str_FecIni, "yyyymmdd") Or _
       Format(moddat_g_str_FecSis, "yyyymmdd") > Format(modctb_str_FecFin, "yyyymmdd")) Then
       MsgBox "Intenta eliminar un documento en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If

   r_dbl_ImpAux = 0
   r_dbl_ImpAux = moddat_gf_ObtieneTipCamDia(2, 2, Format(moddat_g_str_FecSis, "yyyymmdd"), 1)
   If r_dbl_ImpAux = 0 Then
      MsgBox "Tiene que registrar el tipo de cambio sbs del día.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   Call gs_RefrescaGrid(grd_Listad)
   If MsgBox("¿Seguro que desea eliminar el registro seleccionado?" & vbCrLf & _
             "Recuerde sólo es posible eliminar un registro no pagado. ", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
             '"Recuerde sólo es posible eliminar un registro no pagado. Al eliminar la solicitud debe comunicarlo obligatoriamente a Contabilidad de lo contrario causará problemas de conciliación.",
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_CAJCHC_BORRAR ( "
   g_str_Parame = g_str_Parame & "'" & Trim(moddat_g_str_Codigo) & "', " 'CAJCHC_CODCAJ
   g_str_Parame = g_str_Parame & "2, " 'CAJCHC_TIPTAB
   g_str_Parame = g_str_Parame & "NULL, " 'CAJCHC_NUMERO
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo completar la eliminación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   Else
      Call fs_GeneraAsiento(r_str_AsiGen)
      MsgBox "El registro fue eliminado, asiento generado: " & Trim(r_str_AsiGen), vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
   
   Call fs_BuscarCaja
   Call gs_SetFocus(grd_Listad)
End Sub

Public Function fs_ValMod_Aut(p_Codigo As String, p_TipReg As Integer, Optional ByVal p_CodRef As String) As Boolean
   Screen.MousePointer = 11
   fs_ValMod_Aut = True
   '---------------------------------
   'procesado por Compensasion
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT NVL((SELECT DISTINCT COMAUT_CODEST FROM CNTBL_COMAUT A  "
   g_str_Parame = g_str_Parame & "              WHERE A.COMAUT_SITUAC = 1  "
   If p_TipReg = 2 Then
      g_str_Parame = g_str_Parame & "             AND A.COMAUT_TIPOPE = 2 AND TRIM(COMAUT_ORIGEN) = '" & Trim(p_CodRef) & "'"
   End If
   g_str_Parame = g_str_Parame & "                AND A.COMAUT_CODEST IN (1,2,4,5)  "
   g_str_Parame = g_str_Parame & "                AND A.COMAUT_CODOPE = " & CLng(p_Codigo)
   g_str_Parame = g_str_Parame & "                AND ROWNUM = 1 ) "
   g_str_Parame = g_str_Parame & "           ,0) AS CODEST  "
   g_str_Parame = g_str_Parame & "   FROM DUAL  "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then 'ningún registro
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Function
   End If
   
   g_rst_Princi.MoveFirst
   If g_rst_Princi!CODEST <> 0 Then
      MsgBox "El registro se encuentra en el modulo de compensación, no se puede continuar.", vbExclamation, modgen_g_str_NomPlt
      fs_ValMod_Aut = False
      Screen.MousePointer = 0
      Exit Function
   End If
   Screen.MousePointer = 0
End Function

Private Sub cmd_Buscar_Click()
   Call fs_BuscarCaja
   cmb_Empres.Enabled = False
   cmb_Sucurs.Enabled = False
   ipp_FecIni.Enabled = False
   ipp_FecFin.Enabled = False
   chk_Estado.Enabled = False
End Sub

Private Sub cmd_Consul_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CLng(grd_Listad.Text)
      
   moddat_g_int_FlgGrb = 0 'consultar
   
   Call gs_RefrescaGrid(grd_Listad)
   frm_Ctb_EntRen_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Detalle_Click()
Dim r_str_Cadena  As String

   moddat_g_str_Codigo = ""
   moddat_g_str_FecIng = ""
   moddat_g_str_Descri = ""
   moddat_g_str_DesMod = ""
   moddat_g_dbl_MtoPre = 0
   moddat_g_int_Situac = 0
   moddat_g_str_CodMod = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
    
   If grd_Listad.TextMatrix(grd_Listad.Row, 18) = "1" Then
      'ANTICIPOS - TIENE QUE HABER PAGADO
      If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 13)) = "" And CInt(grd_Listad.TextMatrix(grd_Listad.Row, 23)) = 0 Then
         MsgBox "Para adicionar registros, se tiene que realizar el pago en el modulo compensación," & vbCrLf & _
                "por ser un pago anticipo.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
       
   moddat_g_str_Codigo = grd_Listad.TextMatrix(grd_Listad.Row, 0) 'nro entrega a rendir
   moddat_g_str_FecIng = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 1)) 'fecha caja
   r_str_Cadena = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 17)) 'NUMDOC_RESPONSABLE
   moddat_g_str_Descri = r_str_Cadena & " / " & CStr(grd_Listad.TextMatrix(grd_Listad.Row, 3)) 'NOM_RESPONSABLE
   moddat_g_str_DesMod = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 6)) 'moneda
   moddat_g_dbl_MtoPre = grd_Listad.TextMatrix(grd_Listad.Row, 7) 'importe
   moddat_g_int_Situac = grd_Listad.TextMatrix(grd_Listad.Row, 15) 'Flag Proceso
   moddat_g_str_CodMod = CStr(grd_Listad.TextMatrix(grd_Listad.Row, 16)) 'codigo moneda
   
   frm_Ctb_EntRen_03.Show 1
End Sub

Private Sub cmd_Editar_Click()
Dim r_int_Fila   As Integer
    r_int_Fila = 0

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If (grd_Listad.TextMatrix(grd_Listad.Row, 15) = 1) Then 'Flag Proceso
       Call gs_RefrescaGrid(grd_Listad)
       MsgBox "No se pudo editar el registro, la entrega a rendir esta procesada.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
   End If
   Call gs_RefrescaGrid(grd_Listad)
   
   r_int_Fila = grd_Listad.Row
   moddat_g_str_Codigo = CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   moddat_g_int_FlgGrb = 2 'editar
   
   'procesado por Compensasion
    If fs_ValMod_Aut(moddat_g_str_Codigo, 1) = False Then
       Screen.MousePointer = 0
       Exit Sub
    End If
   
   Call gs_UbicaGrid(grd_Listad, r_int_Fila)
   frm_Ctb_EntRen_02.Show 1
   
   Call gs_UbicaGrid(grd_Listad, r_int_Fila)
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

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   cmb_Empres.Enabled = True
   cmb_Sucurs.Enabled = True
   ipp_FecIni.Enabled = True
   ipp_FecFin.Enabled = True
   chk_Estado.Enabled = True
   Call gs_SetFocus(cmb_Empres)
End Sub

Private Sub cmd_Procesar_Click()
Dim r_int_Contad   As Integer
Dim r_bol_Estado   As Boolean
Dim r_bol_EstFil   As Boolean
Dim r_str_CajPrc   As String

   'PROCESADO
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   'VALIDA SI HAY UN REGISTRO SELECCIONADO
   r_bol_Estado = False
   For r_int_Contad = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Contad, 10) = "NO" Then
          If grd_Listad.TextMatrix(r_int_Contad, 14) = "X" Then
             r_bol_Estado = True
             Exit For
          End If
       End If
   Next
   If r_bol_Estado = False Then
      MsgBox "No se han seleccionados registros para procesar.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'VALIDA QUE LOS REGISTRO SELECCIONADOS TENGAN SALDO CERO
   r_bol_Estado = True
   r_bol_EstFil = True
   For r_int_Contad = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Contad, 10) = "NO" And grd_Listad.TextMatrix(r_int_Contad, 14) = "X" Then
          If CDbl(grd_Listad.TextMatrix(r_int_Contad, 9)) <> 0 Then
             r_bol_Estado = False
             'Exit For
          End If
          If CLng(grd_Listad.TextMatrix(r_int_Contad, 24)) > 0 Then
             r_bol_EstFil = False
          End If
       End If
   Next
   If r_bol_Estado = False And r_bol_EstFil = True Then
      MsgBox "Es obligatorio que los registros seleccionados tengan saldo cero (0.00).", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If r_bol_Estado = True And r_bol_EstFil = False Then
      MsgBox "Es obligatorio que todos los documentos sustentados tengan un importe.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If r_bol_Estado = False And r_bol_EstFil = False Then
      MsgBox "Es obligatorio que todos los registros tengan saldo cero (0.00) " & vbCrLf & _
             "y todos los documentos sustentados tengan un importe.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If r_bol_Estado = True And r_bol_EstFil = True Then
      If MsgBox("¿Seguro que desea procesar los registros seleccionados?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      Screen.MousePointer = 11
      For r_int_Contad = 0 To grd_Listad.Rows - 1
          If grd_Listad.TextMatrix(r_int_Contad, 10) = "NO" Then
             If grd_Listad.TextMatrix(r_int_Contad, 14) = "X" Then
                'PROCESANDO REGISTROS
                g_str_Parame = ""
                g_str_Parame = g_str_Parame & " USP_CNTBL_ENTREN_GEN ( "
                g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(r_int_Contad, 0)) & ", "  'CAJDET_CODCAJ
                If Trim(grd_Listad.TextMatrix(r_int_Contad, 18) & "") = "" Then
                   g_str_Parame = g_str_Parame & CLng(0) & ", "  'CAJCHC_TIPPAG
                Else
                   g_str_Parame = g_str_Parame & CLng(grd_Listad.TextMatrix(r_int_Contad, 18)) & ", "  'CAJCHC_TIPPAG
                End If
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
                g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
                g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "

                If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                   MsgBox "La caja " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede procesar.", vbExclamation, modgen_g_str_NomPlt
                   Screen.MousePointer = 0
                   Exit Sub
                End If
                If (g_rst_Genera!RESUL = 1) Then 'insertado
                    r_str_CajPrc = r_str_CajPrc & "- " & Trim(grd_Listad.TextMatrix(r_int_Contad, 0))
                End If
                If (g_rst_Genera!RESUL = 3) Then 'no tiene detalle
                    MsgBox "La caja " & grd_Listad.TextMatrix(r_int_Contad, 0) & " no tiene detalle, no se proceso." & vbCrLf & _
                           "solo se procesaron:" & Trim(r_str_CajPrc), vbExclamation, modgen_g_str_NomPlt
                    Screen.MousePointer = 0
                    Exit Sub
                ElseIf (g_rst_Genera!RESUL = 4) Then 'moneda diferente
                    MsgBox "La caja " & grd_Listad.TextMatrix(r_int_Contad, 0) & " tiene monedas distintas, no se proceso." & vbCrLf & _
                           "solo se procesaron:" & Trim(r_str_CajPrc), vbExclamation, modgen_g_str_NomPlt
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                
             End If
          End If
      Next
        
      MsgBox "Se culminó el proceso de registros seleccionados." & _
             vbCrLf & "Los registros procesados son: " & Trim(r_str_CajPrc), vbInformation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Call fs_BuscarCaja
      Call gs_UbiIniGrid(grd_Listad)
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "NINGUNA"
   cmb_Buscar.AddItem "TIPO PAGO"
   cmb_Buscar.AddItem "RESPONSABLE"
   cmb_Buscar.AddItem "BENEFICIARIO"
   cmb_Buscar.AddItem "PROCESADO"
   
   grd_Listad.ColWidth(0) = 1160 'Nro Caja
   grd_Listad.ColWidth(1) = 1020 'Fecha caja
   grd_Listad.ColWidth(2) = 1230 'tipo pago - 1100
   grd_Listad.ColWidth(3) = 2040 'Responsable
   grd_Listad.ColWidth(4) = 1920 'Beneficiario
   grd_Listad.ColWidth(5) = 1680 'Glosa
   grd_Listad.ColWidth(6) = 660 'Moneda
   grd_Listad.ColWidth(7) = 1080 'Mto Asignado
   grd_Listad.ColWidth(8) = 1000 'Mto Gastado
   grd_Listad.ColWidth(9) = 1000 'Mto Saldo
   grd_Listad.ColWidth(10) = 610 'Procesado
   grd_Listad.ColWidth(11) = 690 'Dias Pendientes
   grd_Listad.ColWidth(12) = 1010  'fecha pago
   grd_Listad.ColWidth(13) = 870 'codigo pago
   grd_Listad.ColWidth(14) = 730 'Selecconar
   grd_Listad.ColWidth(15) = 0 'Flag Proceso
   grd_Listad.ColWidth(16) = 0 'codigo moneda
   grd_Listad.ColWidth(17) = 0 'NUMDOC_RESPONSABLE
   grd_Listad.ColWidth(18) = 0 'CODIGO TIP_PAGO
   grd_Listad.ColWidth(19) = 0 'CajChc_FecCaj
   grd_Listad.ColWidth(20) = 0 'COMPAG_FECPAG
   grd_Listad.ColWidth(21) = 0 'CAJCHC_CODREF_1
   grd_Listad.ColWidth(22) = 0 'CAJCHC_CODREF_2
   grd_Listad.ColWidth(23) = 0 'TOT_FILAS DETALLE
   grd_Listad.ColWidth(24) = 0 'TOT_FILAS DETALLE CERO
      
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_Listad.ColAlignment(10) = flexAlignCenterCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
   grd_Listad.ColAlignment(12) = flexAlignCenterCenter
   grd_Listad.ColAlignment(13) = flexAlignCenterCenter
   grd_Listad.ColAlignment(14) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
Dim r_str_CadAux As String
   cmd_Editar.Enabled = False
   
   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0
   cmb_Empres.ListIndex = 0
   r_str_CadAux = ""
   
   Call moddat_gs_FecSis
   Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
   
   Call moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   r_str_CadAux = DateAdd("m", 1, "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno)
   modctb_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   modctb_str_FecIni = DateAdd("m", -1, modctb_str_FecFin)
   modctb_str_FecIni = "01/" & Format(Month(modctb_str_FecIni), "00") & "/" & Year(modctb_str_FecIni)
   
   ipp_FecIni.Text = modctb_str_FecIni
   ipp_FecFin.Text = modctb_str_FecFin
   
   cmb_Buscar.ListIndex = 0
   cmb_Sucurs.ListIndex = 0
   chk_Estado.Value = 0
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Public Sub fs_BuscarCaja()
Dim r_str_FecIni  As String
Dim r_str_FecFin  As String
Dim r_str_Cadena  As String
Dim r_str_CadAux  As String
Dim r_int_UbiFil  As Integer
Dim r_str_FecPag  As String
Dim r_str_CodPag  As String
Dim r_dbl_Import  As Double

   Screen.MousePointer = 11
   r_str_CadAux = ""
   r_str_Cadena = ""
   If grd_Listad.Row > -1 Then
      r_str_Cadena = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
   r_str_FecIni = Format(ipp_FecIni.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFin.Text, "yyyymmdd")

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.CAJCHC_CODCAJ, A.CAJCHC_FECCAJ, "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_TIPDOC || '-' || A.CAJCHC_NUMDOC NUMDOC_RESPON, TRIM(B.MAEPRV_RAZSOC) NOM_RESPON,  "
   g_str_Parame = g_str_Parame & "        A.CAJCHC_CODMON, TRIM(C.PARDES_DESCRI) MONEDA, A.CAJCHC_FLGPRC, "
   g_str_Parame = g_str_Parame & "        (A.CAJCHC_IMPORT + NVL((SELECT SUM(X.CAJCHC_IMPORT)  "
   g_str_Parame = g_str_Parame & "                             FROM CNTBL_CAJCHC X  "
   g_str_Parame = g_str_Parame & "                            WHERE X.CAJCHC_CODREF_1 = A.CAJCHC_CODCAJ  "
   g_str_Parame = g_str_Parame & "                              AND X.CAJCHC_TIPTAB = 2  "
   g_str_Parame = g_str_Parame & "                              AND X.CAJCHC_SITUAC = 1),0)) AS CAJCHC_IMPORT, "
   g_str_Parame = g_str_Parame & "        (NVL((SELECT SUM(CASE WHEN A.CAJCHC_CODMON = X.CAJDET_CODMON THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0))  "
   g_str_Parame = g_str_Parame & "                              WHEN X.CAJDET_CODMON = 1 THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0)) / Y.TIPCAM_VENTAS  "
   g_str_Parame = g_str_Parame & "                              WHEN X.CAJDET_CODMON = 2 THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0)) * Y.TIPCAM_VENTAS END)  "
   g_str_Parame = g_str_Parame & "                FROM CNTBL_CAJCHC_DET X  "
   g_str_Parame = g_str_Parame & "                LEFT JOIN OPE_TIPCAM Y ON Y.TIPCAM_CODIGO = 3 AND Y.TIPCAM_TIPMON = 2 AND Y.TIPCAM_FECDIA = X.CAJDET_FECEMI  "
   g_str_Parame = g_str_Parame & "               Where x.CajDet_CodCaj = A.CajChc_CodCaj And x.CAJDET_TIPTAB = 2  "
   g_str_Parame = g_str_Parame & "                 AND CAJDET_SITUAC = 1 AND X.CAJDET_TIPCPB NOT IN (7,88)),0)  -  "
   g_str_Parame = g_str_Parame & "         NVL((SELECT SUM(CASE WHEN A.CAJCHC_CODMON = X.CAJDET_CODMON THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0))  "
   g_str_Parame = g_str_Parame & "                              WHEN X.CAJDET_CODMON = 1 THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0)) / Y.TIPCAM_VENTAS  "
   g_str_Parame = g_str_Parame & "                              WHEN X.CAJDET_CODMON = 2 THEN (NVL(X.CAJDET_DEB_PPG1,0) + NVL(X.CAJDET_HAB_PPG1,0)) * Y.TIPCAM_VENTAS END)  "
   g_str_Parame = g_str_Parame & "                FROM CNTBL_CAJCHC_DET X  "
   g_str_Parame = g_str_Parame & "                LEFT JOIN OPE_TIPCAM Y ON Y.TIPCAM_CODIGO = 3 AND Y.TIPCAM_TIPMON = 2 AND Y.TIPCAM_FECDIA = X.CAJDET_FECEMI  "
   g_str_Parame = g_str_Parame & "               Where x.CajDet_CodCaj = A.CajChc_CodCaj And x.CAJDET_TIPTAB = 2  "
   g_str_Parame = g_str_Parame & "                 AND CAJDET_SITUAC = 1 AND X.CAJDET_TIPCPB IN (7,88)),0))  "
   g_str_Parame = g_str_Parame & "        AS MTOGASTADO, A.CAJCHC_DESCRI, A.CAJCHC_NUMOPE,  "
   g_str_Parame = g_str_Parame & "        TRIM(D.MAEPRV_RAZSOC) AS NOM_BENEFI, TRIM(F.PARDES_DESCRI) || DECODE(CAJCHC_CODREF_2,NULL,'',' * ') AS TIPO_PAGO, A.CAJCHC_TIPPAG,  "
   g_str_Parame = g_str_Parame & "        J.COMPAG_FECPAG, J.COMPAG_CODCOM, CAJCHC_CODREF_1, CAJCHC_CODREF_2, "
   g_str_Parame = g_str_Parame & "        (NVL((SELECT SUM(CAJCHC_IMPORT) FROM CNTBL_CAJCHC X  "
   g_str_Parame = g_str_Parame & "              Where x.CajChc_CodCaj = A.CajChc_CodCaj  "
   g_str_Parame = g_str_Parame & "                AND X.CAJCHC_TIPTAB = 4  "
   g_str_Parame = g_str_Parame & "                AND X.CAJCHC_SITUAC = 1),0) +  "
   g_str_Parame = g_str_Parame & "         NVL((SELECT -SUM(CAJCHC_IMPORT) FROM CNTBL_CAJCHC X  "
   g_str_Parame = g_str_Parame & "               Where x.CajChc_CodCaj = A.CajChc_CodCaj  "
   g_str_Parame = g_str_Parame & "                 AND X.CAJCHC_TIPTAB = 5  "
   g_str_Parame = g_str_Parame & "                 AND X.CAJCHC_SITUAC = 1),0)) AS IMPORTE_2,  "
   g_str_Parame = g_str_Parame & "         NVL((SELECT COUNT (*) FROM CNTBL_CAJCHC_DET X  "
   g_str_Parame = g_str_Parame & "               WHERE x.CajDet_CodCaj = A.CajChc_CodCaj And x.CAJDET_TIPTAB = 2  "
   g_str_Parame = g_str_Parame & "                 AND CAJDET_SITUAC = 1),0) NRO_FILDET,  "
   
   g_str_Parame = g_str_Parame & "         NVL((SELECT COUNT(*) FROM CNTBL_CAJCHC_DET X  "
   g_str_Parame = g_str_Parame & "               WHERE x.CajDet_CodCaj = A.CajChc_CodCaj  "
   g_str_Parame = g_str_Parame & "                 AND x.cajdet_tiptab = 2  "
   g_str_Parame = g_str_Parame & "                 AND x.cajdet_situac = 1  "
   g_str_Parame = g_str_Parame & "                 AND (x.CAJDET_Deb_Ppg1 + x.CAJDET_Hab_Ppg1) = 0),0) AS NRO_FILERR  "
                           
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CAJCHC A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = A.CAJCHC_TIPDOC AND B.MAEPRV_NUMDOC = A.CAJCHC_NUMDOC  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 118 AND E.PARDES_CODITE = A.CAJCHC_TIPDOC   " 'tip documento
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 204 AND C.PARDES_CODITE = A.CAJCHC_CODMON  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_MAEPRV D ON D.MAEPRV_TIPDOC = A.CAJCHC_TIPDOC_2 AND D.MAEPRV_NUMDOC = A.CAJCHC_NUMDOC_2  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 138 AND F.PARDES_CODITE = A.CAJCHC_TIPPAG   " 'tipo pago
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMAUT H ON TO_NUMBER(H.COMAUT_CODOPE) = TO_NUMBER(A.CAJCHC_CODCAJ) AND H.COMAUT_TIPOPE = 1 AND H.COMAUT_CODEST NOT IN (3)  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMDET I ON I.COMDET_CODAUT = H.COMAUT_CODAUT AND I.COMDET_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN CNTBL_COMPAG J ON J.COMPAG_CODCOM = I.COMDET_CODCOM AND J.COMPAG_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "                                                                 AND J.COMPAG_FLGCTB = 1  "
   g_str_Parame = g_str_Parame & "   LEFT JOIN MNT_PARDES K ON K.PARDES_CODGRP = 135 AND K.PARDES_CODITE = J.COMPAG_TIPPAG  "
   g_str_Parame = g_str_Parame & "  WHERE A.CAJCHC_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_TIPTAB = 2  "
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_CODREF_1 IS NULL  "
   g_str_Parame = g_str_Parame & "    AND A.CAJCHC_FECCAJ BETWEEN " & r_str_FecIni & " AND " & r_str_FecFin
   
   If chk_Estado.Value = 0 Then
      'solo procesados
      g_str_Parame = g_str_Parame & "    AND A.CAJCHC_FLGPRC  = 0  "
   End If
   If (cmb_Buscar.ListIndex = 1) Then 'tipo pago
       If Len(Trim(txt_Buscar.Text)) > 0 Then
           g_str_Parame = g_str_Parame & "   AND UPPER(TRIM(F.PARDES_DESCRI) || DECODE(CAJCHC_CODREF_2,NULL,'',' * ')) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'"
       End If
   ElseIf (cmb_Buscar.ListIndex = 2) Then 'responsable
       If Len(Trim(txt_Buscar.Text)) > 0 Then
           g_str_Parame = g_str_Parame & "   AND UPPER(TRIM(B.MAEPRV_RAZSOC)) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'"
       End If
   ElseIf (cmb_Buscar.ListIndex = 3) Then 'beneficiario
       If Len(Trim(txt_Buscar.Text)) > 0 Then
           g_str_Parame = g_str_Parame & "   AND UPPER(TRIM(D.MAEPRV_RAZSOC)) LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'"
       End If
   ElseIf (cmb_Buscar.ListIndex = 4) Then 'procesado
       r_str_Cadena = ""
       Select Case UCase(Trim(txt_Buscar.Text))
              Case "S", "SI", "I": r_str_Cadena = "1"
              Case "N", "NO", "O": r_str_Cadena = "0"
       End Select
       If (Len(Trim(r_str_Cadena)) > 0) Then
           g_str_Parame = g_str_Parame & "   AND CAJCHC_FLGPRC = " & r_str_Cadena
       End If
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY CAJCHC_CODCAJ ASC  "
   
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
      grd_Listad.Text = Format(CStr(g_rst_Princi!CajChc_CodCaj), "0000000000")

      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!CajChc_FecCaj)
      
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(CStr(g_rst_Princi!TIPO_PAGO & ""))

      grd_Listad.Col = 3
      grd_Listad.Text = CStr(g_rst_Princi!NOM_RESPON & "")
      
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!NOM_BENEFI & "")
                  
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!CajChc_Descri & "")
      
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!Moneda & "")
      
      grd_Listad.Col = 7 'MTO ASIGNADO
      grd_Listad.Text = Format(g_rst_Princi!CajChc_Import, "###,###,###,##0.00")
      
      grd_Listad.Col = 8 'MTO RENDIDO
      grd_Listad.Text = Format(g_rst_Princi!MTOGASTADO + g_rst_Princi!IMPORTE_2, "###,###,###,##0.00")
               
      grd_Listad.Col = 9 'MTO SALDO
      grd_Listad.Text = Format(g_rst_Princi!CajChc_Import - (g_rst_Princi!MTOGASTADO + g_rst_Princi!IMPORTE_2), "###,###,###,##0.00")
      '--------------------------------------------------------------------------------------------------
      grd_Listad.Col = 10
      grd_Listad.Text = IIf(g_rst_Princi!CAJCHC_FLGPRC = 1, "SI", "NO")
      
      grd_Listad.Col = 11
      If g_rst_Princi!CAJCHC_FLGPRC = 1 Or (g_rst_Princi!CajChc_Import - g_rst_Princi!MTOGASTADO) <= 0 Then
         grd_Listad.Text = 0
      Else
         grd_Listad.Text = DateDiff("D", gf_FormatoFecha(g_rst_Princi!CajChc_FecCaj), moddat_g_str_FecSis)    'DIAS PENDIENTES
      End If
      
      If Trim(g_rst_Princi!COMPAG_FECPAG & "") <> "" Then
         grd_Listad.Col = 12
         grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!COMPAG_FECPAG)
      End If
      If Trim(g_rst_Princi!COMPAG_CODCOM & "") <> "" Then
         grd_Listad.Col = 13
         grd_Listad.Text = Format(g_rst_Princi!COMPAG_CODCOM, "00000000")
      End If
      
      grd_Listad.Col = 15
      grd_Listad.Text = g_rst_Princi!CAJCHC_FLGPRC
      
      grd_Listad.Col = 16
      grd_Listad.Text = g_rst_Princi!CAJCHC_CODMON

      grd_Listad.Col = 17
      grd_Listad.Text = CStr(g_rst_Princi!NUMDOC_RESPON & "")
      
      grd_Listad.Col = 18
      grd_Listad.Text = CStr(g_rst_Princi!cajchc_TipPag & "")
      
      grd_Listad.Col = 19
      grd_Listad.Text = g_rst_Princi!CajChc_FecCaj
      
      If Trim(g_rst_Princi!COMPAG_FECPAG & "") <> "" Then
         grd_Listad.Col = 20
         grd_Listad.Text = g_rst_Princi!COMPAG_FECPAG
      End If
      '------------------------------
      grd_Listad.Col = 21
      grd_Listad.Text = CStr(g_rst_Princi!CAJCHC_CODREF_1 & "")
      grd_Listad.Col = 22
      grd_Listad.Text = CStr(g_rst_Princi!CAJCHC_CODREF_2 & "")
      grd_Listad.Col = 23
      grd_Listad.Text = g_rst_Princi!NRO_FILDET ' DETALLE - TOTAL DE FILAS
      grd_Listad.Col = 24
      grd_Listad.Text = g_rst_Princi!NRO_FILERR ' DETALLE - TOTAL FILAS CERO
      
      g_rst_Princi.MoveNext
   Loop

   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   If r_str_Cadena <> "" Then
      For r_int_UbiFil = 0 To grd_Listad.Rows - 1
          If CLng(CStr(grd_Listad.TextMatrix(r_int_UbiFil, 0))) = CLng(CStr(r_str_Cadena)) Then
             Call gs_UbicaGrid(grd_Listad, r_int_UbiFil)
             grd_Listad.TopRow = r_int_UbiFil
             Exit For
          End If
      Next
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub fs_GeneraAsiento(ByRef p_AsiGen As String)
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_str_AsiGen        As String
Dim r_int_NumAsi        As Integer
Dim r_str_Glosa         As String
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_dbl_TipSbs        As Double
Dim r_str_CtaHab        As String
Dim r_str_CtaDeb        As String
Dim r_str_CadAux        As String
Dim r_str_Respon        As String
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer
Dim r_str_FecPrPgoC     As String
Dim r_str_FecPrPgoL     As String
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 6
   r_str_AsiGen = ""
   r_str_CtaHab = ""
   r_str_CtaDeb = ""
   
   'Inicializa variables
   r_int_NumAsi = 0
   r_str_FecPrPgoC = Format(moddat_g_str_FecSis, "yyyymmdd") 'FECHA ER
   r_str_FecPrPgoL = moddat_g_str_FecSis
      
   'TIPO CAMBIO SBS(2) - VENTA(1)
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, r_str_FecPrPgoC, 1)
      
   r_str_Respon = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 3))
   r_str_Glosa = "ER" & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0)) & "/" & Trim(Mid(r_str_Respon, InStr(1, Trim(r_str_Respon), "-") + 1, Len(Trim(r_str_Respon))))
   r_str_Glosa = Mid(Trim(r_str_Glosa), 1, 60)
   
   r_int_PerMes = modctb_int_PerMes ' Month(r_str_FecPrPgoL)
   r_int_PerAno = modctb_int_PerAno 'Year(r_str_FecPrPgoL)
   
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
   r_str_AsiGen = CStr(r_int_NumAsi)
      
   'Insertar en cabecera
    Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
         r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
                  
   'Insertar en detalle
   r_dbl_MtoSol = 0
   r_dbl_MtoDol = 0
   If grd_Listad.TextMatrix(grd_Listad.Row, 16) = 1 Then 'MONEDA
      'Entrega a rendir Soles:
      r_dbl_MtoSol = CDbl(CStr(grd_Listad.TextMatrix(grd_Listad.Row, 7)))
      r_dbl_MtoDol = Format(CDbl(r_dbl_MtoSol / r_dbl_TipSbs), "###,###,##0.00")
      r_str_CtaDeb = "251419010109"
      r_str_CtaHab = "191807020101"
   Else
      'Entrega a rendir dólares:
      r_dbl_MtoDol = CDbl(CStr(grd_Listad.TextMatrix(grd_Listad.Row, 7)))
      r_dbl_MtoSol = Format(CDbl(r_dbl_MtoDol * r_dbl_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
      r_str_CtaDeb = "252419010109"
      r_str_CtaHab = "192807020101"
   End If
   
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 1, r_str_CtaDeb, CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
                                        
   Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                        r_int_NumAsi, 2, r_str_CtaHab, CDate(r_str_FecPrPgoL), _
                                        r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
   p_AsiGen = r_str_AsiGen
   
   'Actualiza flag de contabilizacion
   r_str_CadAux = ""
   r_str_CadAux = r_str_Origen & "/" & r_int_PerAno & "/" & Format(r_int_PerMes, "00") & "/" & Format(r_int_NumLib, "00") & "/" & r_int_NumAsi
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE CNTBL_CAJCHC "
   g_str_Parame = g_str_Parame & "   SET CAJCHC_DATREV = '" & r_str_CadAux & "' "
   g_str_Parame = g_str_Parame & " WHERE CAJCHC_CODCAJ  = " & CLng(grd_Listad.TextMatrix(grd_Listad.Row, 0))
   g_str_Parame = g_str_Parame & "   AND CAJCHC_TIPTAB  = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
      
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contad        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE REGISTRO DE ENTREGAS A RENDIR"
      .Range(.Cells(2, 2), .Cells(2, 15)).Merge
      .Range(.Cells(2, 2), .Cells(2, 15)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 15)).HorizontalAlignment = xlHAlignCenter

      .Cells(3, 2) = "NRO DE ER"
      .Cells(3, 3) = "FECHA DE ER"
      .Cells(3, 4) = "TIPO DE PAGO"
      .Cells(3, 5) = "RESPONSABLE"
      .Cells(3, 6) = "BENEFICIARIO"
      .Cells(3, 7) = "GLOSA"
      .Cells(3, 8) = "MONEDA"
      .Cells(3, 9) = "MONTO ASIGNADO"
      .Cells(3, 10) = "MONTO RENDIDO"
      .Cells(3, 11) = "MONTO SALDO"
      .Cells(3, 12) = "PROCESADO"
      .Cells(3, 13) = "DIAS PENDIENTE"
      .Cells(3, 14) = "FECHA PAGO"
      .Cells(3, 15) = "CODIGO PAGO"
         
      .Range(.Cells(3, 2), .Cells(3, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 15)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 13 'NRO DE ER
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12 'FECHA DE ER
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15 'tipo de pago
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 40 'RESPONSABLE
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 40 'BENEFICIARIO
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").ColumnWidth = 60 'GLOSA
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 21 'MONEDA
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 18 'IMPORTE ASIGNADO
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 17 'MONTO RENDIDO
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 16 'MONTO SALDO
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("L").ColumnWidth = 12 'PROCESADO
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 15 'DIAS PENDIENTES
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 12 'Fecha PAGO
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 14 'CODIGO PAGO
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(10, 15)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 15)).Font.Size = 11
      
      r_int_NumFil = 2
      For r_int_Contad = 0 To grd_Listad.Rows - 1
          .Cells(r_int_NumFil + 2, 2) = "'" & grd_Listad.TextMatrix(r_int_Contad, 0) 'nro caja
          .Cells(r_int_NumFil + 2, 3) = "'" & grd_Listad.TextMatrix(r_int_Contad, 1) 'fecha de caja
          .Cells(r_int_NumFil + 2, 4) = "'" & grd_Listad.TextMatrix(r_int_Contad, 2) 'tipo de pago
          .Cells(r_int_NumFil + 2, 5) = "'" & grd_Listad.TextMatrix(r_int_Contad, 3) 'responsable
          .Cells(r_int_NumFil + 2, 6) = "'" & grd_Listad.TextMatrix(r_int_Contad, 4) 'beneficiario
          .Cells(r_int_NumFil + 2, 7) = "'" & grd_Listad.TextMatrix(r_int_Contad, 5) 'glosa
          .Cells(r_int_NumFil + 2, 8) = "'" & grd_Listad.TextMatrix(r_int_Contad, 6) 'moneda
          .Cells(r_int_NumFil + 2, 9) = "'" & grd_Listad.TextMatrix(r_int_Contad, 7) 'mto asignado
          .Cells(r_int_NumFil + 2, 10) = "'" & grd_Listad.TextMatrix(r_int_Contad, 8) 'mto rendido
          .Cells(r_int_NumFil + 2, 11) = "'" & grd_Listad.TextMatrix(r_int_Contad, 9) 'mto saldo
          .Cells(r_int_NumFil + 2, 12) = "'" & grd_Listad.TextMatrix(r_int_Contad, 10) 'procesado
          .Cells(r_int_NumFil + 2, 13) = "'" & grd_Listad.TextMatrix(r_int_Contad, 11) 'dias pendientes
          .Cells(r_int_NumFil + 2, 14) = "'" & grd_Listad.TextMatrix(r_int_Contad, 12) 'fecha pago
          .Cells(r_int_NumFil + 2, 15) = "'" & grd_Listad.TextMatrix(r_int_Contad, 13) 'codigo pago
                                                   
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(3, 3), .Cells(3, 15)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 10
      If UCase(grd_Listad.Text) = "NO" Then
         grd_Listad.Col = 14
         If grd_Listad.Text = "X" Then
             grd_Listad.Text = ""
         Else
              grd_Listad.Text = "X"
         End If
      End If
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub chk_Estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(chk_Estado)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub
 
Private Sub txt_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call fs_BuscarCaja
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(*/&%$·#@_.,;:")
   End If
End Sub

Private Sub pnl_Codigo_Click()
   If pnl_Codigo.Tag = "" Then
      pnl_Codigo.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Codigo.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Respon_Click()
   If pnl_Respon.Tag = "" Then
      pnl_Respon.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Respon.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Benefi_Click()
   If pnl_Benefi.Tag = "" Then
      pnl_Benefi.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Benefi.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Glosa_Click()
   If pnl_Glosa.Tag = "" Then
      pnl_Glosa.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Glosa.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Moneda_Click()
   If pnl_Moneda.Tag = "" Then
      pnl_Moneda.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Moneda.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

Private Sub pnl_MtoAsig_Click()
   If pnl_MtoAsig.Tag = "" Then
      pnl_MtoAsig.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 7, "N")
   Else
      pnl_MtoAsig.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 7, "N-")
   End If
End Sub

Private Sub pnl_MtoRen_Click()
   If pnl_MtoRen.Tag = "" Then
      pnl_MtoRen.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 8, "N")
   Else
      pnl_MtoRen.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub pnl_Saldo_Click()
   If pnl_Saldo.Tag = "" Then
      pnl_Saldo.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 9, "N")
   Else
      pnl_Saldo.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 9, "N-")
   End If
End Sub

Private Sub pnl_Proces_Click()
   If pnl_Proces.Tag = "" Then
      pnl_Proces.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 10, "C")
   Else
      pnl_Proces.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 10, "C-")
   End If
End Sub

Private Sub pnl_DiaPte_Click()
   If pnl_DiaPte.Tag = "" Then
      pnl_DiaPte.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 11, "N")
   Else
      pnl_DiaPte.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 11, "N-")
   End If
End Sub

Private Sub pnl_CodPag_Click()
   If pnl_CodPag.Tag = "" Then
      pnl_CodPag.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 13, "C")
   Else
      pnl_CodPag.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 13, "C-")
   End If
End Sub

Private Sub pnl_Selecc_Click()
   If pnl_Selecc.Tag = "" Then
      pnl_Selecc.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 14, "C")
   Else
      pnl_Selecc.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 14, "C-")
   End If
End Sub

Private Sub pnl_FecEnt_Click()
   If pnl_FecEnt.Tag = "" Then
      pnl_FecEnt.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 19, "N")
   Else
      pnl_FecEnt.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 19, "N-")
   End If
End Sub

Private Sub pnl_FecPag_Click()
   If pnl_FecPag.Tag = "" Then
      pnl_FecPag.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 20, "N")
   Else
      pnl_FecPag.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 20, "N-")
   End If
End Sub

Private Sub pnl_TipPag_Click()
   If pnl_TipPag.Tag = "" Then
      pnl_TipPag.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_TipPag.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub




