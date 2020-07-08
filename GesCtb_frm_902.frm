VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_ComCie_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   7590
   ClientLeft      =   4485
   ClientTop       =   2160
   ClientWidth     =   15390
   Icon            =   "GesCtb_frm_902.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7575
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   15420
      _Version        =   65536
      _ExtentX        =   27199
      _ExtentY        =   13361
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
         TabIndex        =   11
         Top             =   60
         Width           =   15285
         _Version        =   65536
         _ExtentX        =   26961
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
            Left            =   630
            TabIndex        =   12
            Top             =   120
            Width           =   2955
            _Version        =   65536
            _ExtentX        =   5212
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Mantenimiento de Comerciales"
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
            Picture         =   "GesCtb_frm_902.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   60
         TabIndex        =   13
         Top             =   795
         Width           =   15285
         _Version        =   65536
         _ExtentX        =   26961
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
         Begin VB.CommandButton cmd_Transferir 
            Enabled         =   0   'False
            Height          =   600
            Left            =   3555
            Picture         =   "GesCtb_frm_902.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Transferir Registros de Mes Anterior"
            Top             =   45
            Width           =   600
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   600
            Left            =   2970
            Picture         =   "GesCtb_frm_902.frx":0BE0
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Borrar Registro"
            Top             =   45
            Width           =   600
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   600
            Left            =   2385
            Picture         =   "GesCtb_frm_902.frx":0EEA
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Editar Registro"
            Top             =   45
            Width           =   600
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   600
            Left            =   1800
            Picture         =   "GesCtb_frm_902.frx":11F4
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Nuevo Registro"
            Top             =   45
            Width           =   600
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   600
            Left            =   1215
            Picture         =   "GesCtb_frm_902.frx":14FE
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   45
            Width           =   600
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   600
            Left            =   630
            Picture         =   "GesCtb_frm_902.frx":1808
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar pantalla"
            Top             =   45
            Width           =   600
         End
         Begin VB.CommandButton cmd_DatCom 
            Height          =   600
            Left            =   45
            Picture         =   "GesCtb_frm_902.frx":1B12
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Datos"
            Top             =   45
            Width           =   600
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14640
            Picture         =   "GesCtb_frm_902.frx":1E1C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   510
         Left            =   60
         TabIndex        =   14
         Top             =   1530
         Width           =   15285
         _Version        =   65536
         _ExtentX        =   26961
         _ExtentY        =   900
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
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   105
            Width           =   1500
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   3645
            TabIndex        =   1
            Top             =   105
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            MinValue        =   "0"
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
         Begin VB.Label Label3 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   135
            TabIndex        =   16
            Top             =   135
            Width           =   795
         End
         Begin VB.Label Label5 
            Caption         =   "Año:"
            Height          =   255
            Left            =   3075
            TabIndex        =   15
            Top             =   135
            Width           =   795
         End
      End
      Begin Threed.SSPanel SSPanel22 
         Height          =   5385
         Left            =   60
         TabIndex        =   17
         Top             =   2115
         Width           =   15285
         _Version        =   65536
         _ExtentX        =   26961
         _ExtentY        =   9499
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   90
            TabIndex        =   18
            Top             =   60
            Width           =   1220
            _Version        =   65536
            _ExtentX        =   2152
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   1260
            TabIndex        =   19
            Top             =   60
            Width           =   1300
            _Version        =   65536
            _ExtentX        =   2293
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DOI"
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   2520
            TabIndex        =   20
            Top             =   60
            Width           =   3810
            _Version        =   65536
            _ExtentX        =   6720
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
         Begin Threed.SSPanel pnl_Tit_FecDes 
            Height          =   285
            Left            =   6300
            TabIndex        =   21
            Top             =   60
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fec. Desemb."
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
         Begin Threed.SSPanel pnl_Tit_TipGar 
            Height          =   285
            Left            =   7455
            TabIndex        =   22
            Top             =   60
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tip. Garantia"
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
         Begin Threed.SSPanel pnl_Tit_MtoGar 
            Height          =   285
            Left            =   12300
            TabIndex        =   23
            Top             =   60
            Width           =   1380
            _Version        =   65536
            _ExtentX        =   2434
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto. Garantia (S/.)"
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
         Begin Threed.SSPanel pnl_Tit_MtoDeu 
            Height          =   285
            Left            =   13665
            TabIndex        =   24
            Top             =   60
            Width           =   1290
            _Version        =   65536
            _ExtentX        =   2275
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto. Saldo (S/.)"
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
         Begin Threed.SSPanel pnl_panel_1 
            Height          =   285
            Left            =   8625
            TabIndex        =   25
            Top             =   60
            Width           =   705
            _Version        =   65536
            _ExtentX        =   1244
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mn. Gar."
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
         Begin Threed.SSPanel pnl_Tit_MtoSal 
            Height          =   285
            Left            =   11190
            TabIndex        =   28
            Top             =   60
            Width           =   1145
            _Version        =   65536
            _ExtentX        =   2020
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto. Saldo"
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
            Height          =   4935
            Left            =   60
            TabIndex        =   3
            Top             =   390
            Width           =   15180
            _ExtentX        =   26776
            _ExtentY        =   8705
            _Version        =   393216
            Rows            =   30
            Cols            =   11
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_MonPre 
            Height          =   285
            Left            =   10470
            TabIndex        =   27
            Top             =   60
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mn. Pres."
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
         Begin Threed.SSPanel pnl_panel_2 
            Height          =   285
            Left            =   9315
            TabIndex        =   26
            Top             =   60
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto. Garantia"
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
Attribute VB_Name = "frm_Mnt_ComCie_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_str_Mes     As String
Dim r_str_Anio    As String

Private Function fs_AnioMes() As String
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COMCIE_PERANO, COMCIE_PERMES "
   g_str_Parame = g_str_Parame & "  FROM (SELECT DISTINCT COMCIE_PERANO, COMCIE_PERMES"
   g_str_Parame = g_str_Parame & "          FROM CRE_COMCIE"
   g_str_Parame = g_str_Parame & "         ORDER BY COMCIE_PERANO DESC, COMCIE_PERMES DESC)"
   g_str_Parame = g_str_Parame & " WHERE ROWNUM < 2 "
   g_str_Parame = g_str_Parame & " ORDER BY COMCIE_PERANO DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   r_str_Mes = Format(g_rst_Princi!COMCIE_PERMES, "00")
   r_str_Anio = g_rst_Princi!COMCIE_PERANO

   fs_AnioMes = r_str_Anio & r_str_Mes
End Function

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_SetFocus(cmb_PerMes)
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = Year(date)
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmd_DatCom.Enabled = Not p_Habilita
   cmd_Limpia.Enabled = p_Habilita
   cmd_ExpExc.Enabled = p_Habilita
   cmd_Agrega.Enabled = p_Habilita
   cmd_Editar.Enabled = p_Habilita
   cmd_Borrar.Enabled = p_Habilita
   cmb_PerMes.Enabled = Not p_Habilita
   ipp_PerAno.Enabled = Not p_Habilita
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1200
   grd_Listad.ColWidth(1) = 1260
   grd_Listad.ColWidth(2) = 3760
   grd_Listad.ColWidth(3) = 1150
   grd_Listad.ColWidth(4) = 1170
   grd_Listad.ColWidth(5) = 680
   grd_Listad.ColWidth(6) = 1155
   grd_Listad.ColWidth(7) = 705
   grd_Listad.ColWidth(8) = 1125
   grd_Listad.ColWidth(9) = 1350
   grd_Listad.ColWidth(10) = 1320
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
      
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
End Sub

Private Sub cmd_Transferir_Click()
   If MsgBox("¿Está seguro de transferir la información del período anterior al período actual?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
    
   g_str_Parame = ""
   g_str_Parame = "SELECT * FROM CRE_COMCIE WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & g_rst_Princi!COMCIE_PERMES & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & g_rst_Princi!COMCIE_PERANO & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY COMCIE_NUMOPE ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
            
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
                  
         'Valida Nivel de Endeudamiento
         If moddat_gf_Consulta_NivelEndeudamiento(g_rst_Princi!COMCIE_TDOCLI, Trim(g_rst_Princi!COMCIE_NDOCLI), r_str_Mes, r_str_Anio, g_rst_Princi!COMCIE_MTOPRE) = True Then
            MsgBox "El Número de Operación " & Left(g_rst_Princi!COMCIE_NUMOPE, 3) & "-" & Mid(g_rst_Princi!COMCIE_NUMOPE, 4, 2) & "-" & Right(g_rst_Princi!COMCIE_NUMOPE, 5) & ", sobrepasa el nivel de endeudamiento permitido según norma, en Créditos Comerciales y/o Cartas Fianza. ", vbExclamation, modgen_g_str_NomPlt
            GoTo Seguir
         End If
            
         With g_rst_Princi
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO CRE_COMCIE ("
            g_str_Parame = g_str_Parame & "  COMCIE_PERMES, COMCIE_PERANO, COMCIE_NUMOPE,"
            g_str_Parame = g_str_Parame & "  COMCIE_CODTIT, COMCIE_FECCIE, COMCIE_TIPCAM, COMCIE_TIPMON, COMCIE_MTOPRE,"
            g_str_Parame = g_str_Parame & "  COMCIE_TOTPRE, COMCIE_PLAMES, COMCIE_CLAPRD, COMCIE_CODPRD, COMCIE_CODSUB,"
            g_str_Parame = g_str_Parame & "  COMCIE_UBIGEO, COMCIE_CODPRY, COMCIE_PRYMCS, COMCIE_ACTECO, COMCIE_CODCIU,"
            g_str_Parame = g_str_Parame & "  COMCIE_SECECO, COMCIE_PERGRA, COMCIE_COSEFE, COMCIE_TASINT, COMCIE_TASMOR,"
            g_str_Parame = g_str_Parame & "  COMCIE_TIPGAR, COMCIE_MONGAR, COMCIE_MTOGAR, COMCIE_FECDES, COMCIE_SITUAC,"
            g_str_Parame = g_str_Parame & "  COMCIE_SALCAP, COMCIE_DIAMOR, COMCIE_SITCRE, COMCIE_TIPCRE, COMCIE_FLGREF,"
            g_str_Parame = g_str_Parame & "  COMCIE_FLGJUD, COMCIE_FLGCAS, COMCIE_CLACRE, COMCIE_CLACLI, COMCIE_CLAALI,"
            g_str_Parame = g_str_Parame & "  COMCIE_PRVGEN, COMCIE_PRVESP, COMCIE_PRVCAM, COMCIE_PRVCIC, COMCIE_PRVADC,"
            g_str_Parame = g_str_Parame & "  COMCIE_FECDEV, COMCIE_DEVVIG, COMCIE_DEVVEN, COMCIE_ACUDVG, COMCIE_ACUDVC,"
            g_str_Parame = g_str_Parame & "  COMCIE_TDOCLI, COMCIE_NDOCLI, COMCIE_CLAPRV, COMCIE_EXPORC, "
            g_str_Parame = g_str_Parame & "  SEGUSUCRE,     SEGFECCRE,     SEGHORCRE,     SEGPLTCRE,     SEGTERCRE,      SEGSUCCRE,"
            g_str_Parame = g_str_Parame & "  SEGUSUACT,     SEGFECACT,     SEGHORACT,     SEGPLTACT,     SEGTERACT,      SEGSUCACT,"
            g_str_Parame = g_str_Parame & "  COMCIE_CAPAMO, COMCIE_FECAMO, COMCIE_APRCRE, COMCIE_ULTVCT, COMCIE_CUOATR,"
            g_str_Parame = g_str_Parame & "  COMCIE_CUOPEN, COMCIE_TIPPAG, COMCIE_CUOPAG, COMCIE_INTDIF, COMCIE_CAPVEN,"
            g_str_Parame = g_str_Parame & "  COMCIE_CAPVIG, COMCIE_VCTANT, COMCIE_PRXVCT, COMCIE_ULTPAG, COMCIE_FVGANT,"
            g_str_Parame = g_str_Parame & "  COMCIE_INTCOM, COMCIE_INTMOR, COMCIE_GASCOB, COMCIE_OTRGAS, COMCIE_UCPPAG,"
            g_str_Parame = g_str_Parame & "  COMCIE_IMOVIG, COMCIE_GCOVIG, COMCIE_OTGVIG, COMCIE_ACUDIF, COMCIE_NUECRE,"
            g_str_Parame = g_str_Parame & "  COMCIE_LINCRE, COMCIE_CREIND, COMCIE_MONCRE, COMCIE_MTOCRE, COMCIE_CODEMP) "
            
            g_str_Parame = g_str_Parame & "VALUES  ('" & r_str_Mes & "', '" & r_str_Anio & "', '" & !COMCIE_NUMOPE & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_CODTIT & "', '" & !comcie_feccie & "', " & !COMCIE_TIPCAM & ","
            g_str_Parame = g_str_Parame & "'" & !COMCIE_TIPMON & "',  " & !COMCIE_MTOPRE & " ,  " & !comcie_totpre & " ,"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_PLAMES & "', '" & !COMCIE_CLAPRD & "', '" & !comcie_codprd & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_CODSUB & "', '" & !COMCIE_UBIGEO & "', '" & !COMCIE_CODPRY & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_PRYMCS & "', '" & !COMCIE_ACTECO & "', '" & !comcie_codciu & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_SECECO & "', '" & !COMCIE_PERGRA & "',  " & !COMCIE_COSEFE & " ,"
            g_str_Parame = g_str_Parame & " " & !COMCIE_TASINT & " ,  " & !comcie_tasmor & " , '" & !comcie_tipgar & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_MONGAR & "',  " & !COMCIE_MTOGAR & " , '" & !COMCIE_FECDES & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_SITUAC & "',  " & !COMCIE_SALCAP & " , '" & !COMCIE_DIAMOR & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_SITCRE & "', '" & !COMCIE_TIPCRE & "', '" & !COMCIE_FLGREF & "',"
            g_str_Parame = g_str_Parame & "'" & !comcie_flgjud & "', '" & !comcie_flgcas & "', '" & !COMCIE_CLACRE & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_CLACLI & "', '" & !COMCIE_CLAALI & "',  " & !COMCIE_PRVGEN & " ,"
            g_str_Parame = g_str_Parame & " " & !COMCIE_PRVESP & " ,  " & !COMCIE_PRVCAM & " ,  " & !COMCIE_PRVCIC & " ,"
            g_str_Parame = g_str_Parame & " " & !COMCIE_PRVADC & " ,  " & !comcie_fecdev & " ,  " & !comcie_devvig & " ,"
            g_str_Parame = g_str_Parame & " " & !comcie_devven & " ,  " & !COMCIE_ACUDVG & " ,  " & !COMCIE_ACUDVC & " ,"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_TDOCLI & "', '" & !COMCIE_NDOCLI & "', '" & !COMCIE_CLAPRV & "',"
            g_str_Parame = g_str_Parame & "'" & !comcie_exporc & "', '" & modgen_g_str_CodUsu & "', '" & Format(date, "yyyymmdd") & "',"
            g_str_Parame = g_str_Parame & "'" & Format(Time, "hhmmss") & "', '" & UCase(App.EXEName) & "', '" & modgen_g_str_NombPC & "',"
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', '', '', '', '', '', '',"
            g_str_Parame = g_str_Parame & " " & !COMCIE_CAPAMO & " , '" & !COMCIE_FECAMO & "', '" & !COMCIE_APRCRE & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_ULTVCT & "', '" & !COMCIE_CUOATR & "', '" & !COMCIE_CUOPEN & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_TIPPAG & "', '" & !COMCIE_CUOPAG & "',  " & !COMCIE_INTDIF & " ,"
            g_str_Parame = g_str_Parame & " " & !COMCIE_CAPVEN & " ,  " & !COMCIE_CAPVIG & " , '" & !COMCIE_VCTANT & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_PRXVCT & "', '" & !COMCIE_ULTPAG & "', '" & !COMCIE_FVGANT & "',"
            g_str_Parame = g_str_Parame & " " & !COMCIE_INTCOM & " ,  " & !COMCIE_INTMOR & " ,  " & !COMCIE_GASCOB & " ,"
            g_str_Parame = g_str_Parame & " " & !COMCIE_OTRGAS & " , '" & !COMCIE_UCPPAG & "',  " & !COMCIE_IMOVIG & " ,"
            g_str_Parame = g_str_Parame & " " & !COMCIE_GCOVIG & " ,  " & !COMCIE_OTGVIG & " ,  " & !COMCIE_ACUDIF & " ,"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_NUECRE & "',  " & !COMCIE_LINCRE & " , '" & !COMCIE_CREIND & "',"
            g_str_Parame = g_str_Parame & "'" & !COMCIE_MONCRE & "',  " & !COMCIE_MTOCRE & " , '" & !COMCIE_CODEMP & ") "
         End With
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            Exit Sub
         End If
Seguir:
         g_rst_Princi.MoveNext
      Loop
      
   Else
      cmd_DatCom.Enabled = False
      MsgBox "Se encontro informacion ya registrada.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call fs_BusDat
   Call fs_Activa(True)
   
   cmd_Transferir.Enabled = False
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Limpia
   Call fs_Activa(False)
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_DatCom_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = "" Then
      MsgBox "Debe seleccionar el Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
      
   Call fs_Activa(True)
   Call fs_BusDat
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(False)
   cmb_PerMes.SetFocus
   cmd_Transferir.Enabled = False
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

Private Sub cmd_Agrega_Click()
   If fs_AnioMes <> ipp_PerAno.Text & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") Then
      MsgBox "No se puede ingresar informacion de ese Periodo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   modsec_g_str_Period = ipp_PerAno.Text & "-" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00")
   moddat_g_int_FlgGrb = 1
   frm_Mnt_ComCie_02.Show 1
   
   Call fs_BusDat
End Sub

Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   modsec_g_str_Period = ipp_PerAno.Text & "-" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00")
   grd_Listad.Col = 0
   moddat_g_str_NumOpe = Left(grd_Listad.Text, 3) & Mid(grd_Listad.Text, 5, 2) & Right(grd_Listad.Text, 5)
   
   If fs_AnioMes <> ipp_PerAno.Text & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") Then
      MsgBox "No se puede modificar informacion del ese Periodo.", vbExclamation, modgen_g_str_NomPlt
      frm_Mnt_ComCie_02.cmd_Editar.Enabled = False
      'Exit Sub
   End If
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgGrb = 2
   frm_Mnt_ComCie_02.Show 1
   Call fs_BusDat
End Sub

Private Sub cmd_Borrar_Click()
   If fs_AnioMes <> ipp_PerAno.Text & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") Then
      MsgBox "No se puede eliminar informacion de ese Periodo.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Estás seguro de eliminar el registro?.", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_NumOpe = Left(grd_Listad.Text, 3) & Mid(grd_Listad.Text, 5, 2) & Right(grd_Listad.Text, 5)
   Call gs_RefrescaGrid(grd_Listad)
   
   g_str_Parame = "DELETE FROM CRE_COMCIE WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_NUMOPE = '" & CStr(moddat_g_str_NumOpe) & "' AND "
   g_str_Parame = g_str_Parame & "COMCIE_SITUAC = " & CInt(1) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   MsgBox "Los datos se eliminaron correctamente.", vbExclamation, modgen_g_str_NomPlt
   Call fs_BusDat
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.COMCIE_NUMOPE, TRIM(B.PRODUC_DESCRI) AS PROUCTO, COMCIE_TDOCLI, COMCIE_NDOCLI, C.DATGEN_RAZSOC, A.COMCIE_CLACRE, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_CLACLI, A.COMCIE_CLAALI, A.COMCIE_CLAPRV, A.COMCIE_PRVGEN, A.COMCIE_PRVESP, A.COMCIE_PRVCAM, A.COMCIE_PRVCIC, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_PRVADC, A.COMCIE_DIAMOR, A.COMCIE_PLAMES, A.COMCIE_PERGRA, A.COMCIE_TASINT, A.COMCIE_TIPGAR, A.COMCIE_MONGAR, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_MTOGAR, A.COMCIE_TIPCAM, A.COMCIE_MTOREA, A.COMCIE_FECDES, A.COMCIE_ULTVCT, A.COMCIE_TIPMON, A.COMCIE_MTOPRE, "
   g_str_Parame = g_str_Parame & "       A.COMCIE_NUECRE, A.COMCIE_LINCRE, A.COMCIE_SALCAP, A.COMCIE_ACUDVG, A.COMCIE_SITUAC, D.TIPCRE_DESCRI, E.PARDES_DESCRI "
   g_str_Parame = g_str_Parame & "  FROM CRE_COMCIE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC B ON B.PRODUC_CODIGO = A.COMCIE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN EMP_DATGEN C ON C.DATGEN_EMPTDO = A.COMCIE_TDOCLI AND C.DATGEN_EMPNDO = A.COMCIE_NDOCLI "
   g_str_Parame = g_str_Parame & " INNER JOIN CTB_TIPCRE D ON D.TIPCRE_CODIGO = A.COMCIE_NUECRE "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = '013' AND E.PARDES_CODITE = A.COMCIE_SITUAC "
   g_str_Parame = g_str_Parame & " WHERE A.COMCIE_PERMES = '" & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & "' "
   g_str_Parame = g_str_Parame & "   AND A.COMCIE_PERANO = '" & CInt(ipp_PerAno.Text) & "' "
   'g_str_Parame = g_str_Parame & "   AND A.COMCIE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY A.COMCIE_NUMOPE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Range(.Cells(1, 36), .Cells(2, 36)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(1, 35), .Cells(2, 36)).Font.Bold = True
      .Range(.Cells(1, 35), .Cells(1, 36)).Merge
      .Range(.Cells(2, 35), .Cells(2, 36)).Merge
      .Cells(1, 35) = "Dpto. de Tecnología e Informática"
      .Cells(2, 35) = "Desarrollo de Sistemas"
      
      .Range(.Cells(5, 1), .Cells(5, 36)).Merge
      .Range(.Cells(5, 1), .Cells(5, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(5, 1), .Cells(5, 1)).Font.Bold = True
      .Range(.Cells(5, 1), .Cells(5, 1)).Font.Underline = xlUnderlineStyleSingle
      .Cells(5, 1) = "CRÉDITOS COMERCIALES"

      .Cells(7, 1) = "ITEM"
      .Cells(7, 2) = "AÑO"
      .Cells(7, 3) = "MES"
      .Cells(7, 4) = "NRO OPERACION"
      .Cells(7, 5) = "PRODUCTO"
      .Cells(7, 6) = "DOC. IDENTIDAD"
      .Cells(7, 7) = "RAZON SOCIAL"
      .Cells(7, 8) = "CLASIFICACION CREDITO (SBS)"
      .Cells(7, 9) = "CLASIFICACION CLIENTE (SBS)"
      .Cells(7, 10) = "CLASIFICACION ALINEADA (SBS)"
      .Cells(7, 11) = "CLASIFICACION PROVISION"
      .Cells(7, 12) = "PROVISION GENERICA"
      .Cells(7, 13) = "PROVISION ESPECIFICA"
      .Cells(7, 14) = "PROVISION RIESGO CAMBIARIO"
      .Cells(7, 15) = "PROVISION PRO-CICLICA"
      .Cells(7, 16) = "PROVISION ADICIONAL"
      .Cells(7, 17) = "DIAS ATRASO"
      .Cells(7, 18) = "PLAZO"
      .Cells(7, 19) = "PERIODO GRACIA"
      .Cells(7, 20) = "TASA INTERES"
      .Cells(7, 21) = "INTERES MORATORIO"
      .Cells(7, 22) = "TIPO DE GARANTIA"
      .Cells(7, 23) = "MONEDA GARANTIA"
      .Cells(7, 24) = "MONTO GARANTIA MN. ORG."
      .Cells(7, 25) = "MONTO GARANTIA SOLES"
      .Cells(7, 26) = "VALOR REALIZACION"
      .Cells(7, 27) = "FECHA DESEMBOLSO"
      .Cells(7, 28) = "FECHA VENCIMIENTO"
      .Cells(7, 29) = "MONEDA PRESTAMO"
      .Cells(7, 30) = "MONTO PRESTAMO MN. ORG."
      .Cells(7, 31) = "MONTO PRESTAMO SOLES"
      .Cells(7, 32) = "TIPO DE CREDITO"
      .Cells(7, 33) = "SALDO LINEA CREDITO"
      .Cells(7, 34) = "SALDO CAPITAL"
      .Cells(7, 35) = "INTERES DEVENGADO"
      .Cells(7, 36) = "SITUACION"

      .Range(.Cells(7, 1), .Cells(7, 36)).Font.Bold = True
      .Range(.Cells(7, 1), .Cells(7, 36)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 36)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(7, 1), .Cells(7, 36)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 36)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 36)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 36)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 36)).Borders(xlInsideVertical).LineStyle = xlContinuous

      .Columns("A").ColumnWidth = 5
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 5
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 14
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 35
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 14
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 40
      .Columns("H").ColumnWidth = 25
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 25
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 25
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 22
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 18
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 18
      .Columns("M").NumberFormat = "###,###,##0.00"
      .Columns("N").ColumnWidth = 25
      .Columns("N").NumberFormat = "###,###,##0.00"
      .Columns("O").ColumnWidth = 20
      .Columns("O").NumberFormat = "###,###,##0.00"
      .Columns("P").ColumnWidth = 18
      .Columns("P").NumberFormat = "###,###,##0.00"
      .Columns("Q").ColumnWidth = 12
      .Columns("Q").HorizontalAlignment = xlHAlignCenter
      .Columns("R").ColumnWidth = 7
      .Columns("R").HorizontalAlignment = xlHAlignCenter
      .Columns("S").ColumnWidth = 14
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      .Columns("T").ColumnWidth = 12
      .Columns("T").NumberFormat = "###,###,##0.00"
      .Columns("U").ColumnWidth = 17
      .Columns("U").NumberFormat = "###,###,##0.00"
      .Columns("V").ColumnWidth = 18
      .Columns("V").HorizontalAlignment = xlHAlignCenter
      .Columns("W").ColumnWidth = 23
      .Columns("W").HorizontalAlignment = xlHAlignCenter
      .Columns("X").ColumnWidth = 23
      .Columns("X").NumberFormat = "###,###,##0.00"
      .Columns("Y").ColumnWidth = 21
      .Columns("Y").NumberFormat = "###,###,##0.00"
      .Columns("Z").ColumnWidth = 17
      .Columns("Z").NumberFormat = "###,###,##0.00"
      .Columns("AA").ColumnWidth = 17
      .Columns("AA").HorizontalAlignment = xlHAlignCenter
      .Columns("AB").ColumnWidth = 17
      .Columns("AB").HorizontalAlignment = xlHAlignCenter
      .Columns("AC").ColumnWidth = 17
      .Columns("AC").HorizontalAlignment = xlHAlignCenter
      .Columns("AD").ColumnWidth = 23
      .Columns("AD").NumberFormat = "###,###,##0.00"
      .Columns("AE").ColumnWidth = 21
      .Columns("AE").NumberFormat = "###,###,##0.00"
      .Columns("AF").ColumnWidth = 30
      .Columns("AF").HorizontalAlignment = xlHAlignCenter
      .Columns("AG").ColumnWidth = 18
      .Columns("AG").NumberFormat = "###,###,##0.00"
      .Columns("AH").ColumnWidth = 14
      .Columns("AH").NumberFormat = "###,###,##0.00"
      .Columns("AI").ColumnWidth = 17
      .Columns("AI").NumberFormat = "###,###,##0.00"
      .Columns("AJ").ColumnWidth = 15
      .Columns("AJ").HorizontalAlignment = xlHAlignCenter
   End With

   g_rst_Princi.MoveFirst
   r_int_ConVer = 8

   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 7
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = CInt(ipp_PerAno.Text)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumOpe(g_rst_Princi!COMCIE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!PROUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CStr(g_rst_Princi!COMCIE_TDOCLI) & "-" & Trim(g_rst_Princi!COMCIE_NDOCLI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!DATGEN_RAZSOC)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = g_rst_Princi!COMCIE_CLACRE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = g_rst_Princi!COMCIE_CLACLI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = g_rst_Princi!COMCIE_CLAALI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = g_rst_Princi!COMCIE_CLAPRV
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!COMCIE_PRVGEN, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!COMCIE_PRVESP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!COMCIE_PRVCAM, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!COMCIE_PRVCIC, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!COMCIE_PRVADC, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = g_rst_Princi!COMCIE_DIAMOR
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = "" & g_rst_Princi!COMCIE_PLAMES
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = g_rst_Princi!COMCIE_PERGRA
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = CStr(g_rst_Princi!COMCIE_TASINT)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = 0
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!comcie_tipgar))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!COMCIE_MONGAR))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!COMCIE_MTOGAR, "###,###,##0.00")
      
      If g_rst_Princi!COMCIE_MONGAR = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!COMCIE_MTOGAR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!COMCIE_MTOGAR * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Format(g_rst_Princi!COMCIE_MTOREA, "###,###,##0.00")
      
      If IsNull(g_rst_Princi!COMCIE_FECDES) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = ""
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = "" & gf_FormatoFecha(CStr(g_rst_Princi!COMCIE_FECDES))
      End If
      If IsNull(g_rst_Princi!COMCIE_ULTVCT) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = ""
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = "" & gf_FormatoFecha(CStr(g_rst_Princi!COMCIE_ULTVCT))
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!COMCIE_TIPMON))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = Format(g_rst_Princi!COMCIE_MTOPRE, "###,###,##0.00")
      If g_rst_Princi!COMCIE_TIPMON = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Format((g_rst_Princi!COMCIE_MTOPRE), "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Format((g_rst_Princi!COMCIE_MTOPRE) * g_rst_Princi!COMCIE_TIPCAM, "###,###,##0.00")
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Trim(g_rst_Princi!TIPCRE_DESCRI)    'COMCIE_NUECRE
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Format((g_rst_Princi!COMCIE_LINCRE), "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 34) = Format((g_rst_Princi!COMCIE_SALCAP), "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 35) = Format((g_rst_Princi!COMCIE_ACUDVG), "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36) = Trim(g_rst_Princi!PARDES_DESCRI)   'COMCIE_SITUAC
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36)).Font.Name = "Arial"
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 36)).Font.Size = 8

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_BusDat()
   Dim r_str_ClaPrd As String '
   Dim r_str_ClaCre As String '
   Dim r_dbl_TipCam As Double '
   
   'fs_Buscar_PorPer = False
   r_str_ClaCre = 0
      
   'Obtener datos
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_COMCIE "
   g_str_Parame = g_str_Parame & " WHERE COMCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " "
   g_str_Parame = g_str_Parame & "   AND COMCIE_PERANO = " & CInt(ipp_PerAno.Text) & " "
   'g_str_Parame = g_str_Parame & "   AND COMCIE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY COMCIE_NUMOPE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
         
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      g_rst_Princi.MoveFirst
            
      Do While Not g_rst_Princi.EOF
         'r_dbl_TipCam = CDbl(g_rst_Princi!COMCIE_TIPCAM)
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = Left(g_rst_Princi!COMCIE_NUMOPE, 3) & "-" & Mid(g_rst_Princi!COMCIE_NUMOPE, 4, 2) & "-" & Right(g_rst_Princi!COMCIE_NUMOPE, 5)
         
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(g_rst_Princi!COMCIE_TDOCLI) & "-" & Trim(g_rst_Princi!COMCIE_NDOCLI)
         
         grd_Listad.Col = 2
         grd_Listad.Text = "" & moddat_gf_Consulta_RazSoc(CStr(g_rst_Princi!COMCIE_TDOCLI), Trim(g_rst_Princi!COMCIE_NDOCLI))
         
         grd_Listad.Col = 3
         If IsNull(g_rst_Princi!COMCIE_FECDES) Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!COMCIE_FECDES))
         End If
                         
         grd_Listad.Col = 4
         grd_Listad.Text = moddat_gf_Consulta_ParDes("066", CStr(g_rst_Princi!comcie_tipgar))
         
         grd_Listad.Col = 5
         If (g_rst_Princi!COMCIE_MONGAR) <> 0 Then
            grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!COMCIE_MONGAR))
         Else
            grd_Listad.Text = " "
         End If
         
         grd_Listad.Col = 6
         If (g_rst_Princi!COMCIE_MTOGAR) <> 0 Then
            grd_Listad.Text = Format(CDbl(g_rst_Princi!COMCIE_MTOGAR), "###,###,##0.00")
         Else
            grd_Listad.Text = " "
         End If
         
         grd_Listad.Col = 7
         grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!COMCIE_TIPMON))
         
         grd_Listad.Col = 8
         grd_Listad.Text = Format(CDbl(g_rst_Princi!COMCIE_SALCAP), "###,###,##0.00")
                 
         Dim l_dbl_MtoGar As Double
         grd_Listad.Col = 9
         If (g_rst_Princi!COMCIE_MTOGAR) <> 0 Then
            If (g_rst_Princi!COMCIE_MONGAR) = 2 Then
               l_dbl_MtoGar = CDbl(g_rst_Princi!COMCIE_MTOGAR * g_rst_Princi!COMCIE_TIPCAM)
               grd_Listad.Text = Format(l_dbl_MtoGar, "###,###,##0.00")
            ElseIf (g_rst_Princi!COMCIE_MONGAR) = 1 Then
               l_dbl_MtoGar = CDbl(g_rst_Princi!COMCIE_MTOGAR)
               grd_Listad.Text = Format(l_dbl_MtoGar, "###,###,##0.00")
            End If
         Else
            grd_Listad.Text = " "
         End If
                
         Dim l_dbl_SalTot As Double
         grd_Listad.Col = 10
         If (g_rst_Princi!COMCIE_TIPMON) = 2 Then
            l_dbl_SalTot = CDbl((g_rst_Princi!COMCIE_SALCAP) * g_rst_Princi!COMCIE_TIPCAM)
            grd_Listad.Text = Format(l_dbl_SalTot, "###,###,##0.00")
         ElseIf (g_rst_Princi!COMCIE_TIPMON) = 1 Then
            l_dbl_SalTot = CDbl(g_rst_Princi!COMCIE_SALCAP)
            grd_Listad.Text = Format(l_dbl_SalTot, "###,###,##0.00")
         End If

         g_rst_Princi.MoveNext
      Loop
      
      'Ordenando por Nombre de Cliente
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
      grd_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Listad)
   Else
      
      '***** proceso para determinar el mes que se debe realizar Transferencia de Registros *****
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT COMCIE_PERANO, COMCIE_PERMES "
      g_str_Parame = g_str_Parame + "  FROM (SELECT DISTINCT COMCIE_PERANO, COMCIE_PERMES"
      g_str_Parame = g_str_Parame + "          FROM CRE_COMCIE"
      g_str_Parame = g_str_Parame + "         ORDER BY COMCIE_PERANO DESC, COMCIE_PERMES DESC)"
      g_str_Parame = g_str_Parame + " WHERE ROWNUM < 2 "
      g_str_Parame = g_str_Parame + " ORDER BY COMCIE_PERANO DESC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      
      r_str_Mes = g_rst_Princi!COMCIE_PERMES
      r_str_Anio = g_rst_Princi!COMCIE_PERANO

      If r_str_Mes = 11 Then
         r_str_Mes = r_str_Mes + 1
      ElseIf r_str_Mes = 12 Then
         r_str_Mes = 1
         r_str_Anio = r_str_Anio + 1
      Else
         r_str_Mes = r_str_Mes + 1
      End If
      
      If CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) = r_str_Mes And CInt(ipp_PerAno.Text) = r_str_Anio Then
         cmd_Transferir.Enabled = True
      Else
         cmd_Transferir.Enabled = False
      End If
      '*******************************************************************************************
      
      cmd_DatCom.Enabled = False
      MsgBox "No se encontraron registros.", vbInformation, modgen_g_str_NomPlt
      
      Exit Sub
   End If
   
   'fs_Buscar_PorPer = True
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Editar_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_PerAno_GotFocus()
   Call gs_SelecTodo(ipp_PerAno)
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_DatCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub
