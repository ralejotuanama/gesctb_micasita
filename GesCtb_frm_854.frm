VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_28 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   6855
   ClientTop       =   1965
   ClientWidth     =   12300
   Icon            =   "GesCtb_frm_854.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8205
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12315
      _Version        =   65536
      _ExtentX        =   21722
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   1440
         Width           =   12240
         _Version        =   65536
         _ExtentX        =   21590
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
         Begin VB.ComboBox cmb_TipDes 
            Height          =   315
            ItemData        =   "GesCtb_frm_854.frx":000C
            Left            =   3120
            List            =   "GesCtb_frm_854.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   180
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   870
            TabIndex        =   0
            Top             =   180
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
         Begin VB.Label Label2 
            Caption         =   "Destino:"
            Height          =   315
            Left            =   2250
            TabIndex        =   12
            Top             =   210
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Año:"
            Height          =   195
            Left            =   330
            TabIndex        =   10
            Top             =   210
            Width           =   330
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   5985
         Left            =   30
         TabIndex        =   11
         Top             =   2130
         Width           =   12240
         _Version        =   65536
         _ExtentX        =   21590
         _ExtentY        =   10557
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
            Height          =   5850
            Left            =   60
            TabIndex        =   13
            Top             =   90
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   10319
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   9
            TabHeight       =   520
            TabCaption(0)   =   "Resumen"
            TabPicture(0)   =   "GesCtb_frm_854.frx":0010
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_LisRes"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Detalle"
            TabPicture(1)   =   "GesCtb_frm_854.frx":002C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_LisDet"
            Tab(1).ControlCount=   1
            Begin Threed.SSPanel SSPanel16 
               Height          =   4425
               Left            =   -74940
               TabIndex        =   14
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   4380
                  Index           =   2
                  Left            =   45
                  TabIndex        =   15
                  Top             =   30
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7726
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   4425
               Left            =   -74940
               TabIndex        =   16
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_inm 
                  Height          =   4380
                  Left            =   45
                  TabIndex        =   17
                  Top             =   30
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7726
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   4425
               Left            =   -74940
               TabIndex        =   18
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   4380
                  Index           =   0
                  Left            =   45
                  TabIndex        =   19
                  Top             =   30
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7726
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin Threed.SSPanel SSPanel33 
               Height          =   4425
               Left            =   -74940
               TabIndex        =   20
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   4380
                  Index           =   3
                  Left            =   45
                  TabIndex        =   21
                  Top             =   30
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7726
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisDet 
               Height          =   5440
               Left            =   -74970
               TabIndex        =   22
               Top             =   360
               Width           =   12090
               _ExtentX        =   21325
               _ExtentY        =   9604
               _Version        =   393216
               Rows            =   5
               Cols            =   13
               BackColorSel    =   32768
               ForeColorSel    =   14737632
               AllowBigSelection=   -1  'True
               FocusRect       =   0
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_LisRes 
               Height          =   5445
               Left            =   30
               TabIndex        =   23
               Top             =   360
               Width           =   12090
               _ExtentX        =   21325
               _ExtentY        =   9604
               _Version        =   393216
               Rows            =   5
               Cols            =   13
               BackColorSel    =   32768
               ForeColorSel    =   14737632
               FocusRect       =   0
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   12240
         _Version        =   65536
         _ExtentX        =   21590
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   270
            Left            =   600
            TabIndex        =   7
            Top             =   120
            Width           =   4365
            _Version        =   65536
            _ExtentX        =   7699
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte de Morosidad - Cartera Atrasada"
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
            Picture         =   "GesCtb_frm_854.frx":0048
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   750
         Width           =   12240
         _Version        =   65536
         _ExtentX        =   21590
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_854.frx":0352
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_854.frx":065C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Procesar informacion"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11640
            Picture         =   "GesCtb_frm_854.frx":0966
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_PerMes     As String
Dim l_str_PerAno     As String
Dim l_str_FecIni     As String
Dim l_str_FecFin     As String
Dim l_str_FecCal     As String
Dim l_str_FecLim     As String
Dim l_dbl_TipCam     As Double
Dim l_str_NomPrd     As String
Dim l_str_FecAno     As String
Dim l_str_NomRpt     As String

Private Sub cmd_Proces_Click()
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   If l_dbl_TipCam = 0 Then
      MsgBox "Debe registrar el tipo de cambio comercial para el día actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Salida)
      Exit Sub
   End If
   If cmb_TipDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el destino del reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDes)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de procesar la información?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_TipRep = 0
   Screen.MousePointer = 11
   l_str_NomPrd = ""
   l_str_FecAno = ""
   l_str_NomRpt = ""
   l_str_NomPrd = cmb_TipDes.Text
   l_str_FecAno = ipp_PerAno.Text
   If cmb_TipDes.ItemData(cmb_TipDes.ListIndex) = 1 Then     'INTERNO
      l_str_NomRpt = "REPORTE_INTERNO"
      moddat_g_int_TipRep = 1
      Call fs_Obtiene_Detalle("")
      Call fs_Obtiene_Resumem("")
      moddat_g_int_OrdAct = 30
   ElseIf cmb_TipDes.ItemData(cmb_TipDes.ListIndex) = 2 Then 'OPIC
      l_str_NomRpt = "REPORTE_OPIC"
      moddat_g_int_TipRep = 2
      Call fs_Obtiene_Detalle_OPIC
      Call fs_Obtiene_Resumen_OPIC
      moddat_g_int_OrdAct = 0
   ElseIf cmb_TipDes.ItemData(cmb_TipDes.ListIndex) = 3 Then 'MIVIVIENDA
      l_str_NomRpt = "REPORTE_MIVIVIENDA"
      moddat_g_int_TipRep = 3
      Call fs_Obtiene_Detalle_MVV
      Call fs_Obtiene_Resumen_MVV
      moddat_g_int_OrdAct = 0
   ElseIf cmb_TipDes.ItemData(cmb_TipDes.ListIndex) = 4 Then 'MICASITA
      l_str_NomRpt = "REPORTE_MICASITA"
      moddat_g_int_TipRep = 4
      Call fs_Obtiene_Detalle_MC
      Call fs_Obtiene_Resumen_MC
      moddat_g_int_OrdAct = 0
   ElseIf cmb_TipDes.ItemData(cmb_TipDes.ListIndex) = 5 Then 'MICROEMPRESARIO
      l_str_NomRpt = "REPORTE_MICROEMPRESARIO"
      moddat_g_int_TipRep = 5
      Call fs_Obtiene_Detalle("MICROEMPRESARIO")
      Call fs_Obtiene_Resumem("MICROEMPRESARIO")
      moddat_g_int_OrdAct = 30
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   If grd_LisDet.Rows = 2 Then
      MsgBox "Debe procesar la informacion para poder exportarla.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Proces)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
'   If moddat_g_int_TipRep <> 1 And moddat_g_int_TipRep <> 2 And moddat_g_int_TipRep <> 5 Then
'      Call fs_GenExc2 'MIVIENDA, MICASITA
'   Else
'      Call fs_GenExc 'INTERNO, OPIC, MICROEMPRESARIO
'   End If
   If moddat_g_int_TipRep <> 1 And moddat_g_int_TipRep <> 2 And moddat_g_int_TipRep <> 5 Then
      Call fs_GenExc2 'MIVIENDA, MICASITA
   Else
      Call fs_GenExc 'INTERNO, OPIC, MICROEMPRESARIO
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Busca_UltCierre
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(ipp_PerAno)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   ipp_PerAno.Text = Year(date)
   
   cmb_TipDes.Clear
   cmb_TipDes.AddItem "INTERNO"
   cmb_TipDes.ItemData(cmb_TipDes.NewIndex) = 1
   cmb_TipDes.AddItem "OPIC"
   cmb_TipDes.ItemData(cmb_TipDes.NewIndex) = 2
   cmb_TipDes.AddItem "MIVIVIENDA"
   cmb_TipDes.ItemData(cmb_TipDes.NewIndex) = 3
   cmb_TipDes.AddItem "MICASITA"
   cmb_TipDes.ItemData(cmb_TipDes.NewIndex) = 4
   cmb_TipDes.AddItem "MICROEMPRESARIO"
   cmb_TipDes.ItemData(cmb_TipDes.NewIndex) = 5
   cmb_TipDes.ListIndex = -1
   
   'DETALLE
   grd_LisDet.Redraw = False
   Call gs_LimpiaGrid(grd_LisDet)
   grd_LisDet.SelectionMode = flexSelectionByRow
   grd_LisDet.FocusRect = flexFocusNone
   grd_LisDet.HighLight = flexHighlightAlways
   grd_LisDet.ColWidth(0) = 2000
   grd_LisDet.ColWidth(1) = 800
   grd_LisDet.ColWidth(2) = 800
   grd_LisDet.ColWidth(3) = 800
   grd_LisDet.ColWidth(4) = 800
   grd_LisDet.ColWidth(5) = 800
   grd_LisDet.ColWidth(6) = 800
   grd_LisDet.ColWidth(7) = 800
   grd_LisDet.ColWidth(8) = 800
   grd_LisDet.ColWidth(9) = 800
   grd_LisDet.ColWidth(10) = 800
   grd_LisDet.ColWidth(11) = 800
   grd_LisDet.ColWidth(12) = 800
   grd_LisDet.ColAlignment(0) = flexAlignLeftCenter
   grd_LisDet.ColAlignment(1) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(2) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(3) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(4) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(5) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(6) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(7) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(8) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(9) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(10) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(11) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(12) = flexAlignCenterCenter
      
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Col = 0:  grd_LisDet.Text = Space(12) & "PRODUCTO"
   grd_LisDet.Col = 1:  grd_LisDet.Text = "ENE"
   grd_LisDet.Col = 2:  grd_LisDet.Text = "FEB"
   grd_LisDet.Col = 3:  grd_LisDet.Text = "MAR"
   grd_LisDet.Col = 4:  grd_LisDet.Text = "ABR"
   grd_LisDet.Col = 5:  grd_LisDet.Text = "MAY"
   grd_LisDet.Col = 6:  grd_LisDet.Text = "JUN"
   grd_LisDet.Col = 7:  grd_LisDet.Text = "JUL"
   grd_LisDet.Col = 8:  grd_LisDet.Text = "AGO"
   grd_LisDet.Col = 9:  grd_LisDet.Text = "SET"
   grd_LisDet.Col = 10: grd_LisDet.Text = "OCT"
   grd_LisDet.Col = 11: grd_LisDet.Text = "NOV"
   grd_LisDet.Col = 12: grd_LisDet.Text = "DIC"
   
   With grd_LisDet
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 1
      .Row = 1
      .ColSel = 1
   End With
   grd_LisDet.Redraw = True
   
   'RESUMEN
   grd_LisRes.Redraw = False
   Call gs_LimpiaGrid(grd_LisRes)
   grd_LisRes.SelectionMode = flexSelectionByRow
   grd_LisRes.FocusRect = flexFocusNone
   grd_LisRes.HighLight = flexHighlightAlways
   grd_LisRes.ColWidth(0) = 2000
   grd_LisRes.ColWidth(1) = 800
   grd_LisRes.ColWidth(2) = 800
   grd_LisRes.ColWidth(3) = 800
   grd_LisRes.ColWidth(4) = 800
   grd_LisRes.ColWidth(5) = 800
   grd_LisRes.ColWidth(6) = 800
   grd_LisRes.ColWidth(7) = 800
   grd_LisRes.ColWidth(8) = 800
   grd_LisRes.ColWidth(9) = 800
   grd_LisRes.ColWidth(10) = 800
   grd_LisRes.ColWidth(11) = 800
   grd_LisRes.ColWidth(12) = 800
   grd_LisRes.ColAlignment(0) = flexAlignLeftCenter
   grd_LisRes.ColAlignment(1) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(2) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(3) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(4) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(5) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(6) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(7) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(8) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(9) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(10) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(11) = flexAlignCenterCenter
   grd_LisRes.ColAlignment(12) = flexAlignCenterCenter
      
   grd_LisRes.Rows = grd_LisRes.Rows + 1
   grd_LisRes.Col = 0:  grd_LisRes.Text = Space(12) & "PRODUCTO"
   grd_LisRes.Col = 1:  grd_LisRes.Text = "ENE"
   grd_LisRes.Col = 2:  grd_LisRes.Text = "FEB"
   grd_LisRes.Col = 3:  grd_LisRes.Text = "MAR"
   grd_LisRes.Col = 4:  grd_LisRes.Text = "ABR"
   grd_LisRes.Col = 5:  grd_LisRes.Text = "MAY"
   grd_LisRes.Col = 6:  grd_LisRes.Text = "JUN"
   grd_LisRes.Col = 7:  grd_LisRes.Text = "JUL"
   grd_LisRes.Col = 8:  grd_LisRes.Text = "AGO"
   grd_LisRes.Col = 9:  grd_LisRes.Text = "SET"
   grd_LisRes.Col = 10: grd_LisRes.Text = "OCT"
   grd_LisRes.Col = 11: grd_LisRes.Text = "NOV"
   grd_LisRes.Col = 12: grd_LisRes.Text = "DIC"
   
   With grd_LisRes
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 1
      .Row = 1
      .ColSel = 1
   End With
   grd_LisRes.Redraw = True
End Sub

Private Sub fs_Busca_UltCierre()
Dim r_str_PerMes  As String
Dim r_str_PerAno  As String
   
   'Obtiene ultimo cierre procesado
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & " FROM (SELECT DISTINCT HIPCIE_PERANO, HIPCIE_PERMES "
   g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & "        ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
   g_str_Parame = g_str_Parame & " WHERE ROWNUM < 2 "
   g_str_Parame = g_str_Parame & " ORDER BY HIPCIE_PERANO DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   l_str_PerMes = g_rst_Princi!HIPCIE_PERMES
   l_str_PerAno = g_rst_Princi!HIPCIE_PERANO
   
   If CInt(l_str_PerMes) = 12 Then
      r_str_PerMes = 1
      r_str_PerAno = CInt(l_str_PerAno) + 1
   Else
      r_str_PerMes = CInt(l_str_PerMes) + 1
      r_str_PerAno = CInt(l_str_PerAno)
   End If
   
   l_str_PerMes = r_str_PerMes
   l_str_PerAno = r_str_PerAno
   l_str_FecIni = Format(r_str_PerAno, "0000") & Format(r_str_PerMes, "00") & "01"
   l_str_FecFin = Format(r_str_PerAno, "0000") & Format(r_str_PerMes, "00") & ff_Ultimo_Dia_Mes(CInt(r_str_PerMes), CInt(r_str_PerAno))
   l_str_FecLim = Format(CDate(moddat_g_str_FecSis), "yyyymmdd")
   If Month(CDate(moddat_g_str_FecSis)) <> CInt(l_str_PerMes) Then
      l_str_FecLim = Format(DateAdd("d", -30, Format(Mid(l_str_FecFin, 7, 2) & "/" & Mid(l_str_FecFin, 5, 2) & "/" & Mid(l_str_FecFin, 1, 4), "DD/MM/YYYY")), "YYYYMMDD")
   End If
      
   'Obtiene ultimo tipo de cambio ingresado
   If (CDbl(Format(moddat_g_str_FecSis, "yyyymmdd")) <= CDbl(l_str_FecFin)) Then
       l_dbl_TipCam = modprc_gf_TipoCambio(1, 1, 2, Format(CDate(moddat_g_str_FecSis), "yyyymmdd"))
       l_str_FecCal = Format(CDate(moddat_g_str_FecSis), "yyyymmdd")
   Else
       l_dbl_TipCam = modprc_gf_TipoCambio(1, 1, 2, l_str_FecFin)
       l_str_FecCal = l_str_FecFin
   End If
End Sub

Private Sub fs_Obtiene_Detalle(p_SubProd As String)
Dim r_rst_MorDia     As ADODB.Recordset
Dim r_dbl_MorMiv     As Double
Dim r_dbl_MorMic     As Double
Dim r_dbl_MorCme     As Double
Dim r_dbl_MorCrc     As Double
Dim r_dbl_ProTot     As Double
Dim r_int_Mor_01     As Integer
Dim r_int_Mor_02     As Integer
Dim r_dbl_CanTot     As Double
Dim r_dbl_CarPor     As Double
Dim r_int_ConFil     As Integer
Dim r_int_ConCol     As Integer
      
   'DETALLE
   grd_LisDet.Rows = 20
   grd_LisDet.TextMatrix(1, 0) = "CRC-PBP"
   grd_LisDet.TextMatrix(2, 0) = "MICASITA"
   grd_LisDet.TextMatrix(3, 0) = "CME"
   grd_LisDet.TextMatrix(4, 0) = "MIVIVIENDA"
   grd_LisDet.TextMatrix(5, 0) = "MICASAMAS"
   grd_LisDet.TextMatrix(6, 0) = "BBP"
   grd_LisDet.TextMatrix(7, 0) = "TECHO PROPIO"
   
   grd_LisDet.TextMatrix(8, 0) = "TOTAL"
   grd_LisDet.TextMatrix(9, 0) = "CAR"
   grd_LisDet.TextMatrix(10, 0) = "TOTAL CANTIDAD"
   grd_LisDet.TextMatrix(11, 0) = "Morosos > a 90 días"
   grd_LisDet.TextMatrix(12, 0) = "Morosos > a 30 días"
   grd_LisDet.TextMatrix(13, 0) = "DESEMBOLSADOS"
   grd_LisDet.TextMatrix(14, 0) = "CANCELADOS"
   grd_LisDet.TextMatrix(15, 0) = "TRANSFERIDOS"
   grd_LisDet.TextMatrix(16, 0) = "% Morosos > a 30 días"
   grd_LisDet.TextMatrix(17, 0) = "% Morosos > a 60 días"
   grd_LisDet.TextMatrix(18, 0) = "% Morosos > a 90 días"
   grd_LisDet.TextMatrix(19, 0) = "% Morosos > a 120 días"
   
   'Obtiene la morosidad de los periodos actuales
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_02("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & " 0 , "
   g_str_Parame = g_str_Parame & l_str_FecLim & " , "
   g_str_Parame = g_str_Parame & l_str_FecCal & " , "
   g_str_Parame = g_str_Parame & l_str_FecIni & " , "
   g_str_Parame = g_str_Parame & l_dbl_TipCam & " , "
   g_str_Parame = g_str_Parame & " '" & p_SubProd & "', " 'sub-producto
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
      
   Select Case CInt(l_str_PerMes)
          Case 1: r_int_ConCol = 1
          Case 2: r_int_ConCol = 2
          Case 3: r_int_ConCol = 3
          Case 4: r_int_ConCol = 4
          Case 5: r_int_ConCol = 5
          Case 6: r_int_ConCol = 6
          Case 7: r_int_ConCol = 7
          Case 8: r_int_ConCol = 8
          Case 9: r_int_ConCol = 9
          Case 10: r_int_ConCol = 10
          Case 11: r_int_ConCol = 11
          Case 12: r_int_ConCol = 12
   End Select
   
   r_rst_MorDia.MoveFirst
   Do While Not r_rst_MorDia.EOF
      grd_LisDet.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00") 'CRC-PBP
      grd_LisDet.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM02), 0, r_rst_MorDia!RPT_VALNUM02), "##0.00") 'MICASITA
      grd_LisDet.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00") 'CME
      grd_LisDet.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM04), 0, r_rst_MorDia!RPT_VALNUM04), "##0.00") 'MIVIVIENDA
      
      grd_LisDet.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM05), 0, r_rst_MorDia!RPT_VALNUM05), "##0.00") 'MICASAMAS
      grd_LisDet.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM06), 0, r_rst_MorDia!RPT_VALNUM06), "##0.00") 'BBP
      grd_LisDet.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM07), 0, r_rst_MorDia!RPT_VALNUM07), "##0.00") 'TECHO PROPIO
            
      grd_LisDet.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
      grd_LisDet.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
      grd_LisDet.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
      grd_LisDet.TextMatrix(11, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
      grd_LisDet.TextMatrix(12, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
      grd_LisDet.TextMatrix(13, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
      grd_LisDet.TextMatrix(14, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
      grd_LisDet.TextMatrix(15, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS
      
      grd_LisDet.TextMatrix(16, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM16), 0, r_rst_MorDia!RPT_VALNUM16), "##0.00")   'Morosos > a 30 días
      grd_LisDet.TextMatrix(17, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM17), 0, r_rst_MorDia!RPT_VALNUM17), "##0.00")   'Morosos > a 60 días
      grd_LisDet.TextMatrix(18, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM18), 0, r_rst_MorDia!RPT_VALNUM18), "##0.00")   'Morosos > a 90 días
      grd_LisDet.TextMatrix(19, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM19), 0, r_rst_MorDia!RPT_VALNUM19), "##0.00")   'Morosos > a 120 días
      
      r_rst_MorDia.MoveNext
   Loop
      
   'Obtiene la morosidad de los periodos cerrados del año
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_01("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & " 0 , "
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
      
   r_rst_MorDia.MoveFirst
   r_int_ConCol = 1
   Do While Not r_rst_MorDia.EOF
      If (l_str_PerAno & Format(l_str_PerMes, "00") <> r_rst_MorDia!RPT_PERANO & Format(r_rst_MorDia!RPT_PERMES, "00")) Then
          grd_LisDet.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00")    'CRC-PBP
          grd_LisDet.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM02), 0, r_rst_MorDia!RPT_VALNUM02), "##0.00")    'MICASITA
          grd_LisDet.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00")    'CME
          grd_LisDet.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM04), 0, r_rst_MorDia!RPT_VALNUM04), "##0.00")    'MIVIVIENDA
          
          grd_LisDet.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM05), 0, r_rst_MorDia!RPT_VALNUM05), "##0.00")    'MICASAMAS
          grd_LisDet.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM06), 0, r_rst_MorDia!RPT_VALNUM06), "##0.00")    'BBP
          grd_LisDet.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM07), 0, r_rst_MorDia!RPT_VALNUM07), "##0.00")    'TECHO PROPIO
          
          grd_LisDet.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00")    'TOTAL
          grd_LisDet.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00")    'CAR
          grd_LisDet.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0")       'TOTAL CANTIDAD
          grd_LisDet.TextMatrix(11, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0")       'Morosos > a 90 días
          grd_LisDet.TextMatrix(12, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0")       'Morosos > a 0 días
          grd_LisDet.TextMatrix(13, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0")      'DESEMBOLSADOS
          grd_LisDet.TextMatrix(14, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0")      'CANCELADOS
          grd_LisDet.TextMatrix(15, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0")      'TRANSFERIDOS
          grd_LisDet.TextMatrix(16, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM16), 0, r_rst_MorDia!RPT_VALNUM16), "##0.00")   'Morosos > a 30 días
          grd_LisDet.TextMatrix(17, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM17), 0, r_rst_MorDia!RPT_VALNUM17), "##0.00")   'Morosos > a 60 días
          grd_LisDet.TextMatrix(18, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM18), 0, r_rst_MorDia!RPT_VALNUM18), "##0.00")   'Morosos > a 90 días
          grd_LisDet.TextMatrix(19, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM19), 0, r_rst_MorDia!RPT_VALNUM19), "##0.00")   'Morosos > a 120 días
      End If
      r_int_ConCol = r_int_ConCol + 1
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing
   
   Call gs_UbiIniGrid(grd_LisDet)
End Sub

Private Sub fs_Obtiene_Resumem(p_SubProd As String)
Dim r_rst_MorDia     As ADODB.Recordset
Dim r_dbl_MorMiv     As Double
Dim r_dbl_MorMic     As Double
Dim r_dbl_MorCme     As Double
Dim r_dbl_MorCrc     As Double
Dim r_dbl_ProTot     As Double
Dim r_int_Mor_01     As Integer
Dim r_int_Mor_02     As Integer
Dim r_dbl_CanTot     As Double
Dim r_dbl_CarPor     As Double
Dim r_int_ConFil     As Integer
Dim r_int_ConCol     As Integer
         
   'RESUMEN
   grd_LisRes.Rows = 17
   grd_LisRes.TextMatrix(1, 0) = "MIVIVIENDA"
   grd_LisRes.TextMatrix(2, 0) = "TECHO PROPIO"
   grd_LisRes.TextMatrix(3, 0) = "MICASITA"
   grd_LisRes.TextMatrix(4, 0) = "OTROS"
   
   grd_LisRes.TextMatrix(5, 0) = "TOTAL"
   grd_LisRes.TextMatrix(6, 0) = "CAR"
   grd_LisRes.TextMatrix(7, 0) = "TOTAL CANTIDAD"
   grd_LisRes.TextMatrix(8, 0) = "Morosos > a 90 días"
   grd_LisRes.TextMatrix(9, 0) = "Morosos > a 30 días"
   grd_LisRes.TextMatrix(10, 0) = "DESEMBOLSADOS"
   grd_LisRes.TextMatrix(11, 0) = "CANCELADOS"
   grd_LisRes.TextMatrix(12, 0) = "TRANSFERIDOS"
   grd_LisRes.TextMatrix(13, 0) = "% Morosos > a 30 días"
   grd_LisRes.TextMatrix(14, 0) = "% Morosos > a 60 días"
   grd_LisRes.TextMatrix(15, 0) = "% Morosos > a 90 días"
   grd_LisRes.TextMatrix(16, 0) = "% Morosos > a 120 días"
   
   'Obtiene la morosidad de los periodos actuales
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_03("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & " 0 , "
   g_str_Parame = g_str_Parame & l_str_FecLim & " , "
   g_str_Parame = g_str_Parame & l_str_FecCal & " , "
   g_str_Parame = g_str_Parame & l_str_FecIni & " , "
   g_str_Parame = g_str_Parame & l_dbl_TipCam & " , "
   g_str_Parame = g_str_Parame & " '" & p_SubProd & "', " 'sub-producto
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
      
   Select Case CInt(l_str_PerMes)
          Case 1: r_int_ConCol = 1
          Case 2: r_int_ConCol = 2
          Case 3: r_int_ConCol = 3
          Case 4: r_int_ConCol = 4
          Case 5: r_int_ConCol = 5
          Case 6: r_int_ConCol = 6
          Case 7: r_int_ConCol = 7
          Case 8: r_int_ConCol = 8
          Case 9: r_int_ConCol = 9
          Case 10: r_int_ConCol = 10
          Case 11: r_int_ConCol = 11
          Case 12: r_int_ConCol = 12
   End Select
   
   r_rst_MorDia.MoveFirst
   Do While Not r_rst_MorDia.EOF
      grd_LisRes.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00") 'MIVIVIENDA
      grd_LisRes.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM02), 0, r_rst_MorDia!RPT_VALNUM02), "##0.00") 'MICASITA
      grd_LisRes.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00") 'TECHO PROPIO
      grd_LisRes.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM04), 0, r_rst_MorDia!RPT_VALNUM04), "##0.00") 'OTROS
      
      grd_LisRes.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
      grd_LisRes.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
      grd_LisRes.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
      grd_LisRes.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
      grd_LisRes.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
      grd_LisRes.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
      grd_LisRes.TextMatrix(11, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
      grd_LisRes.TextMatrix(12, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS
      
      grd_LisRes.TextMatrix(13, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM16), 0, r_rst_MorDia!RPT_VALNUM16), "##0.00")   'Morosos > a 30 días
      grd_LisRes.TextMatrix(14, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM17), 0, r_rst_MorDia!RPT_VALNUM17), "##0.00")   'Morosos > a 60 días
      grd_LisRes.TextMatrix(15, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM18), 0, r_rst_MorDia!RPT_VALNUM18), "##0.00")   'Morosos > a 90 días
      grd_LisRes.TextMatrix(16, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM19), 0, r_rst_MorDia!RPT_VALNUM19), "##0.00")   'Morosos > a 120 días
      
      r_rst_MorDia.MoveNext
   Loop
      
   'Obtiene la morosidad de los periodos cerrados del año
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_04("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & " 0 , "
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
      
   r_rst_MorDia.MoveFirst
   r_int_ConCol = 1
   Do While Not r_rst_MorDia.EOF
      If (l_str_PerAno & Format(l_str_PerMes, "00") <> r_rst_MorDia!RPT_PERANO & Format(r_rst_MorDia!RPT_PERMES, "00")) Then
          grd_LisRes.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00")    'MIVIVIENDA
          grd_LisRes.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM02), 0, r_rst_MorDia!RPT_VALNUM02), "##0.00")    'TECHO PROPIO
          grd_LisRes.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00")    'MICASITA
          grd_LisRes.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM04), 0, r_rst_MorDia!RPT_VALNUM04), "##0.00")    'OTROS
          
          grd_LisRes.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00")    'TOTAL
          grd_LisRes.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00")    'CAR
          grd_LisRes.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0")       'TOTAL CANTIDAD
          grd_LisRes.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0")       'Morosos > a 90 días
          grd_LisRes.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0")       'Morosos > a 0 días
          grd_LisRes.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0")      'DESEMBOLSADOS
          grd_LisRes.TextMatrix(11, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0")      'CANCELADOS
          grd_LisRes.TextMatrix(12, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0")      'TRANSFERIDOS
          grd_LisRes.TextMatrix(13, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM16), 0, r_rst_MorDia!RPT_VALNUM16), "##0.00")   'Morosos > a 30 días
          grd_LisRes.TextMatrix(14, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM17), 0, r_rst_MorDia!RPT_VALNUM17), "##0.00")   'Morosos > a 60 días
          grd_LisRes.TextMatrix(15, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM18), 0, r_rst_MorDia!RPT_VALNUM18), "##0.00")   'Morosos > a 90 días
          grd_LisRes.TextMatrix(16, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM19), 0, r_rst_MorDia!RPT_VALNUM19), "##0.00")   'Morosos > a 120 días
      End If
      r_int_ConCol = r_int_ConCol + 1
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing
   
   Call gs_UbiIniGrid(grd_LisRes)
End Sub

Private Sub fs_Obtiene_Detalle_OPIC()
Dim r_rst_MorDia     As ADODB.Recordset
Dim r_int_ConFil     As Integer
Dim r_int_ConCol     As Integer

   'DETALLE
   grd_LisDet.Rows = 16
   grd_LisDet.TextMatrix(1, 0) = "CRC-PBP"
   grd_LisDet.TextMatrix(2, 0) = "MICASITA"
   grd_LisDet.TextMatrix(3, 0) = "CME"
   grd_LisDet.TextMatrix(4, 0) = "MIVIVIENDA"
   grd_LisDet.TextMatrix(5, 0) = "MICASAMAS"
   grd_LisDet.TextMatrix(6, 0) = "BBP"
   grd_LisDet.TextMatrix(7, 0) = "TECHO PROPIO"
   
   grd_LisDet.TextMatrix(8, 0) = "TOTAL"
   grd_LisDet.TextMatrix(9, 0) = "CAR"
   grd_LisDet.TextMatrix(10, 0) = "TOTAL CANTIDAD"
   grd_LisDet.TextMatrix(11, 0) = "Morosos > a 90 días"
   grd_LisDet.TextMatrix(12, 0) = "Morosos > a 0 días"
   grd_LisDet.TextMatrix(13, 0) = "DESEMBOLSADOS"
   grd_LisDet.TextMatrix(14, 0) = "CANCELADOS"
   grd_LisDet.TextMatrix(15, 0) = "TRANSFERIDOS"
   
   'Obtiene la morosidad de los periodos cerrados del año
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_01("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & " 0, "
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
      
   r_rst_MorDia.MoveFirst
   r_int_ConCol = 1
   Do While Not r_rst_MorDia.EOF
      grd_LisDet.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00") 'CRC-PBP
      grd_LisDet.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM02), 0, r_rst_MorDia!RPT_VALNUM02), "##0.00") 'MICASITA
      grd_LisDet.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00") 'CME
      grd_LisDet.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM04), 0, r_rst_MorDia!RPT_VALNUM04), "##0.00") 'MIVIVIENDA
      
      grd_LisDet.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM05), 0, r_rst_MorDia!RPT_VALNUM05), "##0.00") 'MICASAMAS
      grd_LisDet.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM06), 0, r_rst_MorDia!RPT_VALNUM06), "##0.00") 'BBP
      grd_LisDet.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM07), 0, r_rst_MorDia!RPT_VALNUM07), "##0.00") 'TECHO PROPIO
      
      grd_LisDet.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
      grd_LisDet.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
      grd_LisDet.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
      grd_LisDet.TextMatrix(11, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
      grd_LisDet.TextMatrix(12, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
      grd_LisDet.TextMatrix(13, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
      grd_LisDet.TextMatrix(14, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
      grd_LisDet.TextMatrix(15, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS
            
      r_int_ConCol = r_int_ConCol + 1
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing
      
   Call gs_UbiIniGrid(grd_LisDet)
End Sub

Private Sub fs_Obtiene_Resumen_OPIC()
Dim r_rst_MorDia     As ADODB.Recordset
Dim r_int_ConFil     As Integer
Dim r_int_ConCol     As Integer

   'RESUMEN
   grd_LisRes.Rows = 13
   grd_LisRes.TextMatrix(1, 0) = "MIVIVIENDA"
   grd_LisRes.TextMatrix(2, 0) = "TECHO PROPIO"
   grd_LisRes.TextMatrix(3, 0) = "MICASITA"
   grd_LisRes.TextMatrix(4, 0) = "OTROS"

   grd_LisRes.TextMatrix(5, 0) = "TOTAL"
   grd_LisRes.TextMatrix(6, 0) = "CAR"
   grd_LisRes.TextMatrix(7, 0) = "TOTAL CANTIDAD"
   grd_LisRes.TextMatrix(8, 0) = "Morosos > a 90 días"
   grd_LisRes.TextMatrix(9, 0) = "Morosos > a 0 días"
   grd_LisRes.TextMatrix(10, 0) = "DESEMBOLSADOS"
   grd_LisRes.TextMatrix(11, 0) = "CANCELADOS"
   grd_LisRes.TextMatrix(12, 0) = "TRANSFERIDOS"
   
   'Obtiene la morosidad de los periodos cerrados del año
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_04("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & " 0, "
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
      
   r_rst_MorDia.MoveFirst
   r_int_ConCol = 1
   Do While Not r_rst_MorDia.EOF
      grd_LisRes.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00") 'MIVIVIENDA
      grd_LisRes.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM02), 0, r_rst_MorDia!RPT_VALNUM02), "##0.00") 'TECHO PROPIO
      grd_LisRes.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00") 'MICASITA
      grd_LisRes.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM04), 0, r_rst_MorDia!RPT_VALNUM04), "##0.00") 'OTROS
      
      grd_LisRes.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
      grd_LisRes.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
      grd_LisRes.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
      grd_LisRes.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
      grd_LisRes.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
      grd_LisRes.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
      grd_LisRes.TextMatrix(11, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
      grd_LisRes.TextMatrix(12, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS
            
      r_int_ConCol = r_int_ConCol + 1
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing
      
   Call gs_UbiIniGrid(grd_LisRes)
End Sub

Private Sub fs_Obtiene_Detalle_MVV()
Dim r_rst_MorDia     As ADODB.Recordset
Dim r_dbl_MorMiv     As Double
Dim r_dbl_MorMic     As Double
Dim r_dbl_MorCme     As Double
Dim r_dbl_MorCrc     As Double
Dim r_dbl_ProTot     As Double
Dim r_int_Mor_01     As Integer
Dim r_int_Mor_02     As Integer
Dim r_dbl_CanTot     As Double
Dim r_dbl_CarPor     As Double
Dim r_int_ConFil     As Integer
Dim r_int_ConCol     As Integer
   
   'DETALLE
   grd_LisDet.Rows = 11
   grd_LisDet.TextMatrix(1, 0) = "MIVIVIENDA"
   grd_LisDet.TextMatrix(2, 0) = "CME"
   grd_LisDet.TextMatrix(3, 0) = "TOTAL"
   grd_LisDet.TextMatrix(4, 0) = "CAR"
   grd_LisDet.TextMatrix(5, 0) = "TOTAL CANTIDAD"
   grd_LisDet.TextMatrix(6, 0) = "Morosos > a 90 días"
   grd_LisDet.TextMatrix(7, 0) = "Morosos > a 0 días"
   grd_LisDet.TextMatrix(8, 0) = "DESEMBOLSADOS"
   grd_LisDet.TextMatrix(9, 0) = "CANCELADOS"
   grd_LisDet.TextMatrix(10, 0) = "TRANSFERIDOS"
   
   'Obtiene la morosidad de los periodos actuales
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_02("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & 0 & " , "
   g_str_Parame = g_str_Parame & l_str_FecLim & " , "
   g_str_Parame = g_str_Parame & l_str_FecCal & " , "
   g_str_Parame = g_str_Parame & l_str_FecIni & " , "
   g_str_Parame = g_str_Parame & l_dbl_TipCam & " , "
   g_str_Parame = g_str_Parame & " '', " 'sub-producto
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
   
   Select Case CInt(l_str_PerMes)
          Case 1: r_int_ConCol = 1
          Case 2: r_int_ConCol = 2
          Case 3: r_int_ConCol = 3
          Case 4: r_int_ConCol = 4
          Case 5: r_int_ConCol = 5
          Case 6: r_int_ConCol = 6
          Case 7: r_int_ConCol = 7
          Case 8: r_int_ConCol = 8
          Case 9: r_int_ConCol = 9
          Case 10: r_int_ConCol = 10
          Case 11: r_int_ConCol = 11
          Case 12: r_int_ConCol = 12
   End Select
   
   r_rst_MorDia.MoveFirst
   Do While Not r_rst_MorDia.EOF
      grd_LisDet.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM04), 0, r_rst_MorDia!RPT_VALNUM04), "##0.00") 'MIVIVIENDA
      grd_LisDet.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00") 'CME
      grd_LisDet.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
      grd_LisDet.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
      grd_LisDet.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
      grd_LisDet.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
      grd_LisDet.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
      grd_LisDet.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
      grd_LisDet.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
      grd_LisDet.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS
            
'      grd_LisDet.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00") 'MIVIVIENDA
'      grd_LisDet.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00") 'CME
'      grd_LisDet.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM05), 0, r_rst_MorDia!RPT_VALNUM05), "##0.00") 'TOTAL
'      grd_LisDet.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM06), 0, r_rst_MorDia!RPT_VALNUM06), "##0.00") 'CAR
'      grd_LisDet.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM07), 0, r_rst_MorDia!RPT_VALNUM07), "##0") 'TOTAL CANTIDAD
'      grd_LisDet.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0") 'Morosos > a 90 días
'      grd_LisDet.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0") 'Morosos > a 0 días
'      grd_LisDet.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'DESEMBOLSADOS
'      grd_LisDet.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'CANCELADOS
'      grd_LisDet.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'TRANSFERIDOS
            
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing
   
   'Obtiene la morosidad de los periodos cerrados del año
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_01("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & " 0 , "
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
   
   r_rst_MorDia.MoveFirst
   r_int_ConCol = 1
   Do While Not r_rst_MorDia.EOF
      If (l_str_PerAno & Format(l_str_PerMes, "00") <> r_rst_MorDia!RPT_PERANO & Format(r_rst_MorDia!RPT_PERMES, "00")) Then
          grd_LisDet.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM04), 0, r_rst_MorDia!RPT_VALNUM04), "##0.00") 'MIVIVIENDA
          grd_LisDet.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00") 'CME
          grd_LisDet.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
          grd_LisDet.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
          grd_LisDet.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
          grd_LisDet.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
          grd_LisDet.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
          grd_LisDet.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
          grd_LisDet.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
          grd_LisDet.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS

      End If
      r_int_ConCol = r_int_ConCol + 1
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing
         
   Call gs_UbiIniGrid(grd_LisDet)
End Sub

Private Sub fs_Obtiene_Resumen_MVV()
Dim r_rst_MorDia     As ADODB.Recordset
Dim r_dbl_MorMiv     As Double
Dim r_dbl_MorMic     As Double
Dim r_dbl_MorCme     As Double
Dim r_dbl_MorCrc     As Double
Dim r_dbl_ProTot     As Double
Dim r_int_Mor_01     As Integer
Dim r_int_Mor_02     As Integer
Dim r_dbl_CanTot     As Double
Dim r_dbl_CarPor     As Double
Dim r_int_ConFil     As Integer
Dim r_int_ConCol     As Integer
      
   'RESUMEN
   grd_LisRes.Rows = 10
   grd_LisRes.TextMatrix(1, 0) = "MIVIVIENDA"
   grd_LisRes.TextMatrix(2, 0) = "TOTAL"
   grd_LisRes.TextMatrix(3, 0) = "CAR"
   grd_LisRes.TextMatrix(4, 0) = "TOTAL CANTIDAD"
   grd_LisRes.TextMatrix(5, 0) = "Morosos > a 90 días"
   grd_LisRes.TextMatrix(6, 0) = "Morosos > a 0 días"
   grd_LisRes.TextMatrix(7, 0) = "DESEMBOLSADOS"
   grd_LisRes.TextMatrix(8, 0) = "CANCELADOS"
   grd_LisRes.TextMatrix(9, 0) = "TRANSFERIDOS"
   
   'Obtiene la morosidad de los periodos actuales
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_03("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & 0 & " , "
   g_str_Parame = g_str_Parame & l_str_FecLim & " , "
   g_str_Parame = g_str_Parame & l_str_FecCal & " , "
   g_str_Parame = g_str_Parame & l_str_FecIni & " , "
   g_str_Parame = g_str_Parame & l_dbl_TipCam & " , "
   g_str_Parame = g_str_Parame & " '', " 'sub-producto
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
   
   Select Case CInt(l_str_PerMes)
          Case 1: r_int_ConCol = 1
          Case 2: r_int_ConCol = 2
          Case 3: r_int_ConCol = 3
          Case 4: r_int_ConCol = 4
          Case 5: r_int_ConCol = 5
          Case 6: r_int_ConCol = 6
          Case 7: r_int_ConCol = 7
          Case 8: r_int_ConCol = 8
          Case 9: r_int_ConCol = 9
          Case 10: r_int_ConCol = 10
          Case 11: r_int_ConCol = 11
          Case 12: r_int_ConCol = 12
   End Select
   
   r_rst_MorDia.MoveFirst
   Do While Not r_rst_MorDia.EOF
      grd_LisRes.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00") 'MIVIVIENDA
      grd_LisRes.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
      grd_LisRes.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
      grd_LisRes.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
      grd_LisRes.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
      grd_LisRes.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
      grd_LisRes.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
      grd_LisRes.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
      grd_LisRes.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS
            
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing
   
   'Obtiene la morosidad de los periodos cerrados del año
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_04("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & " 0 , "
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
   
   r_rst_MorDia.MoveFirst
   r_int_ConCol = 1
   Do While Not r_rst_MorDia.EOF
      If (l_str_PerAno & Format(l_str_PerMes, "00") <> r_rst_MorDia!RPT_PERANO & Format(r_rst_MorDia!RPT_PERMES, "00")) Then
          grd_LisRes.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00") 'MIVIVIENDA
          grd_LisRes.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
          grd_LisRes.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
          grd_LisRes.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
          grd_LisRes.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
          grd_LisRes.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
          grd_LisRes.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
          grd_LisRes.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
          grd_LisRes.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS

      End If
      r_int_ConCol = r_int_ConCol + 1
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing
         
   Call gs_UbiIniGrid(grd_LisRes)
End Sub

Private Sub fs_Obtiene_Detalle_MC()
Dim r_rst_MorDia     As ADODB.Recordset
Dim r_dbl_MorMiv     As Double
Dim r_dbl_MorMic     As Double
Dim r_dbl_MorCme     As Double
Dim r_dbl_MorCrc     As Double
Dim r_dbl_ProTot     As Double
Dim r_int_Mor_01     As Integer
Dim r_int_Mor_02     As Integer
Dim r_dbl_CanTot     As Double
Dim r_dbl_CarPor     As Double
Dim r_int_ConFil     As Integer
Dim r_int_ConCol     As Integer

   'DETALLE
   grd_LisDet.Rows = 11
   grd_LisDet.TextMatrix(1, 0) = "MICASITA"
   grd_LisDet.TextMatrix(2, 0) = "CRC-PBP"
   grd_LisDet.TextMatrix(3, 0) = "TOTAL"
   grd_LisDet.TextMatrix(4, 0) = "CAR"
   grd_LisDet.TextMatrix(5, 0) = "TOTAL CANTIDAD"
   grd_LisDet.TextMatrix(6, 0) = "Morosos > a 90 días"
   grd_LisDet.TextMatrix(7, 0) = "Morosos > a 0 días"
   grd_LisDet.TextMatrix(8, 0) = "DESEMBOLSADOS"
   grd_LisDet.TextMatrix(9, 0) = "CANCELADOS"
   grd_LisDet.TextMatrix(10, 0) = "TRANSFERIDOS"
      
   'Obtiene la morosidad de los periodos actuales
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_02("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & 0 & " , "
   g_str_Parame = g_str_Parame & l_str_FecLim & " , "
   g_str_Parame = g_str_Parame & l_str_FecCal & " , "
   g_str_Parame = g_str_Parame & l_str_FecIni & " , "
   g_str_Parame = g_str_Parame & l_dbl_TipCam & " , "
   g_str_Parame = g_str_Parame & " '', " 'sub-producto
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
      
   Select Case CInt(l_str_PerMes)
          Case 1: r_int_ConCol = 1
          Case 2: r_int_ConCol = 2
          Case 3: r_int_ConCol = 3
          Case 4: r_int_ConCol = 4
          Case 5: r_int_ConCol = 5
          Case 6: r_int_ConCol = 6
          Case 7: r_int_ConCol = 7
          Case 8: r_int_ConCol = 8
          Case 9: r_int_ConCol = 9
          Case 10: r_int_ConCol = 10
          Case 11: r_int_ConCol = 11
          Case 12: r_int_ConCol = 12
   End Select

   r_rst_MorDia.MoveFirst
   Do While Not r_rst_MorDia.EOF
      grd_LisDet.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM02), 0, r_rst_MorDia!RPT_VALNUM02), "##0.00") 'MICASITA
      grd_LisDet.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00") 'MICASITA-PBP
      
      grd_LisDet.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
      grd_LisDet.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
      grd_LisDet.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
      grd_LisDet.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
      grd_LisDet.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
      grd_LisDet.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
      grd_LisDet.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
      grd_LisDet.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS
            
'      grd_LisDet.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM05), 0, r_rst_MorDia!RPT_VALNUM05), "##0.00") 'TOTAL
'      grd_LisDet.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM06), 0, r_rst_MorDia!RPT_VALNUM06), "##0.00") 'CAR
'      grd_LisDet.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM07), 0, r_rst_MorDia!RPT_VALNUM07), "##0") 'TOTAL CANTIDAD
'      grd_LisDet.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0") 'Morosos > a 90 días
'      grd_LisDet.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0") 'Morosos > a 0 días
'      grd_LisDet.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'DESEMBOLSADOS
'      grd_LisDet.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'CANCELADOS
'      grd_LisDet.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'TRANSFERIDOS
            
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing

   'Obtiene la morosidad de los periodos cerrados del año
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_01("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & " 0 , "
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
   
   r_rst_MorDia.MoveFirst
   r_int_ConCol = 1
   Do While Not r_rst_MorDia.EOF
      If (l_str_PerAno & Format(l_str_PerMes, "00") <> r_rst_MorDia!RPT_PERANO & Format(r_rst_MorDia!RPT_PERMES, "00")) Then
          grd_LisDet.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM02), 0, r_rst_MorDia!RPT_VALNUM02), "##0.00") 'MICASITA
          grd_LisDet.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM01), 0, r_rst_MorDia!RPT_VALNUM01), "##0.00") 'MICASITA-PBP
          
          grd_LisDet.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
          grd_LisDet.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
          grd_LisDet.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
          grd_LisDet.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
          grd_LisDet.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
          grd_LisDet.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
          grd_LisDet.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
          grd_LisDet.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS
          
'          grd_LisDet.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM04), 0, r_rst_MorDia!RPT_VALNUM04), "##0.00") 'MICASITA-PBP
'          grd_LisDet.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM05), 0, r_rst_MorDia!RPT_VALNUM05), "##0.00") 'TOTAL
'          grd_LisDet.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM06), 0, r_rst_MorDia!RPT_VALNUM06), "##0.00") 'CAR
'          grd_LisDet.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM07), 0, r_rst_MorDia!RPT_VALNUM07), "##0") 'TOTAL CANTIDAD
'          grd_LisDet.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0") 'Morosos > a 90 días
'          grd_LisDet.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0") 'Morosos > a 0 días
'          grd_LisDet.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'DESEMBOLSADOS
'          grd_LisDet.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'CANCELADOS
'          grd_LisDet.TextMatrix(10, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'TRANSFERIDOS

      End If
      r_int_ConCol = r_int_ConCol + 1
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing
   
   Call gs_UbiIniGrid(grd_LisDet)
End Sub

Private Sub fs_Obtiene_Resumen_MC()
Dim r_rst_MorDia     As ADODB.Recordset
Dim r_dbl_MorMiv     As Double
Dim r_dbl_MorMic     As Double
Dim r_dbl_MorCme     As Double
Dim r_dbl_MorCrc     As Double
Dim r_dbl_ProTot     As Double
Dim r_int_Mor_01     As Integer
Dim r_int_Mor_02     As Integer
Dim r_dbl_CanTot     As Double
Dim r_dbl_CarPor     As Double
Dim r_int_ConFil     As Integer
Dim r_int_ConCol     As Integer
   
   'RESUMEN
   grd_LisRes.Rows = 10
   grd_LisRes.TextMatrix(1, 0) = "MICASITA"
   grd_LisRes.TextMatrix(2, 0) = "TOTAL"
   grd_LisRes.TextMatrix(3, 0) = "CAR"
   grd_LisRes.TextMatrix(4, 0) = "TOTAL CANTIDAD"
   grd_LisRes.TextMatrix(5, 0) = "Morosos > a 90 días"
   grd_LisRes.TextMatrix(6, 0) = "Morosos > a 0 días"
   grd_LisRes.TextMatrix(7, 0) = "DESEMBOLSADOS"
   grd_LisRes.TextMatrix(8, 0) = "CANCELADOS"
   grd_LisRes.TextMatrix(9, 0) = "TRANSFERIDOS"
   
   'Obtiene la morosidad de los periodos actuales
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_03("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & 0 & " , "
   g_str_Parame = g_str_Parame & l_str_FecLim & " , "
   g_str_Parame = g_str_Parame & l_str_FecCal & " , "
   g_str_Parame = g_str_Parame & l_str_FecIni & " , "
   g_str_Parame = g_str_Parame & l_dbl_TipCam & " , "
   g_str_Parame = g_str_Parame & " '', " 'sub-producto
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
      
   Select Case CInt(l_str_PerMes)
          Case 1: r_int_ConCol = 1
          Case 2: r_int_ConCol = 2
          Case 3: r_int_ConCol = 3
          Case 4: r_int_ConCol = 4
          Case 5: r_int_ConCol = 5
          Case 6: r_int_ConCol = 6
          Case 7: r_int_ConCol = 7
          Case 8: r_int_ConCol = 8
          Case 9: r_int_ConCol = 9
          Case 10: r_int_ConCol = 10
          Case 11: r_int_ConCol = 11
          Case 12: r_int_ConCol = 12
   End Select

   r_rst_MorDia.MoveFirst
   Do While Not r_rst_MorDia.EOF
      grd_LisRes.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00") 'MICASITA
      
      grd_LisRes.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
      grd_LisRes.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
      grd_LisRes.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
      grd_LisRes.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
      grd_LisRes.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
      grd_LisRes.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
      grd_LisRes.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
      grd_LisRes.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing

   'Obtiene la morosidad de los periodos cerrados del año
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_MORACAR_04("
   g_str_Parame = g_str_Parame & ipp_PerAno.Text & " , "
   g_str_Parame = g_str_Parame & " 0 , "
   g_str_Parame = g_str_Parame & moddat_g_int_TipRep & " , "
   g_str_Parame = g_str_Parame & "'" & l_str_NomRpt & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT * FROM RPT_TABLA_TEMP A "
   g_str_Parame = g_str_Parame & "   WHERE A.RPT_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "     AND A.RPT_TERCRE = " & "'" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_USUCRE = " & "'" & modgen_g_str_CodUsu & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_NOMBRE = '" & l_str_NomRpt & "' "
   g_str_Parame = g_str_Parame & "     AND A.RPT_MONEDA = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_MorDia, 3) Then
      Exit Sub
   End If
   
   r_rst_MorDia.MoveFirst
   r_int_ConCol = 1
   Do While Not r_rst_MorDia.EOF
      If (l_str_PerAno & Format(l_str_PerMes, "00") <> r_rst_MorDia!RPT_PERANO & Format(r_rst_MorDia!RPT_PERMES, "00")) Then
          grd_LisRes.TextMatrix(1, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM03), 0, r_rst_MorDia!RPT_VALNUM03), "##0.00") 'MICASITA
          
          grd_LisRes.TextMatrix(2, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM08), 0, r_rst_MorDia!RPT_VALNUM08), "##0.00") 'TOTAL
          grd_LisRes.TextMatrix(3, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM09), 0, r_rst_MorDia!RPT_VALNUM09), "##0.00") 'CAR
          grd_LisRes.TextMatrix(4, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM10), 0, r_rst_MorDia!RPT_VALNUM10), "##0") 'TOTAL CANTIDAD
          grd_LisRes.TextMatrix(5, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM11), 0, r_rst_MorDia!RPT_VALNUM11), "##0") 'Morosos > a 90 días
          grd_LisRes.TextMatrix(6, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM12), 0, r_rst_MorDia!RPT_VALNUM12), "##0") 'Morosos > a 0 días
          grd_LisRes.TextMatrix(7, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM13), 0, r_rst_MorDia!RPT_VALNUM13), "##0") 'DESEMBOLSADOS
          grd_LisRes.TextMatrix(8, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM14), 0, r_rst_MorDia!RPT_VALNUM14), "##0") 'CANCELADOS
          grd_LisRes.TextMatrix(9, r_int_ConCol) = Format(IIf(IsNull(r_rst_MorDia!RPT_VALNUM15), 0, r_rst_MorDia!RPT_VALNUM15), "##0") 'TRANSFERIDOS
      End If
      r_int_ConCol = r_int_ConCol + 1
      r_rst_MorDia.MoveNext
   Loop
   
   r_rst_MorDia.Close
   Set r_rst_MorDia = Nothing
   
   Call gs_UbiIniGrid(grd_LisRes)
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel     As Excel.Application
Dim r_int_NroFil    As Integer
Dim r_int_Contador  As Integer
Dim r_str_Cadena    As String
Dim r_int_TotFil    As Integer
Dim r_int_NumAux1    As Integer
Dim r_int_NumAux2    As Integer

   'Preparando Cabecera de Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 5
         
   'FORMATEA TITULO Y COLUMNAS GENERAL
   r_obj_Excel.Sheets(1).Name = "RESUMIDO"
   With r_obj_Excel.Sheets(1)
      .Cells(3, 3) = "CARTERA ATRASADA: Vencido + Judicial/Total Crédito ó Tasa de Morosidad (%)"
      .Range(.Cells(3, 3), .Cells(3, 15)).Merge
      .Range("C3:O3").HorizontalAlignment = xlHAlignCenter
      .Range("C3:O3").Font.Bold = True
      .Range(.Cells(3, 3), .Cells(3, 15)).Font.Size = 14
      .Range("C4:O4").HorizontalAlignment = xlHAlignCenter
      .Range("C4:O4").Font.Bold = True
      .Cells(4, 3) = l_str_FecAno & "  -  " & Trim(l_str_NomPrd)
      .Range(.Cells(4, 3), .Cells(4, 15)).Merge
      
      .Range("C5:O25").Font.Name = "Arial Narrow"
      .Range("C5:O25").Font.Size = 10
      
      .Cells(r_int_NroFil, 3) = "PRODUCTO"
      .Cells(r_int_NroFil, 4) = "ENE"
      .Cells(r_int_NroFil, 5) = "FEB"
      .Cells(r_int_NroFil, 6) = "MAR"
      .Cells(r_int_NroFil, 7) = "ABR"
      .Cells(r_int_NroFil, 8) = "MAY"
      .Cells(r_int_NroFil, 9) = "JUN"
      .Cells(r_int_NroFil, 10) = "JUL"
      .Cells(r_int_NroFil, 11) = "AGO"
      .Cells(r_int_NroFil, 12) = "SET"
      .Cells(r_int_NroFil, 13) = "OCT"
      .Cells(r_int_NroFil, 14) = "NOV"
      .Cells(r_int_NroFil, 15) = "DIC"
            
      .Cells(r_int_NroFil + 1, 3) = "MIVIVIENDA"
      .Cells(r_int_NroFil + 2, 3) = "TECHO PROPIO"
      .Cells(r_int_NroFil + 3, 3) = "MICASITA"
      .Cells(r_int_NroFil + 4, 3) = "OTROS"
      .Cells(r_int_NroFil + 5, 3) = "TOTAL"
      .Cells(r_int_NroFil + 6, 3) = "CAR"
      .Cells(r_int_NroFil + 7, 3) = "TOTAL CANTIDAD"
      .Cells(r_int_NroFil + 8, 3) = "Morosos > a 90 días"
      .Cells(r_int_NroFil + 9, 3) = "Morosos > a 30 días"
      .Cells(r_int_NroFil + 10, 3) = "DESEMBOLSADOS"
      .Cells(r_int_NroFil + 11, 3) = "CANCELADOS"
      .Cells(r_int_NroFil + 12, 3) = "TRANSFERIDOS"
      '-------"REPORTE_INTERNO"--------"REPORTE_MICROEMPRESARIO"-------
      If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
         .Cells(r_int_NroFil + 13, 3) = "% Morosos > a 30 días"
         .Cells(r_int_NroFil + 14, 3) = "% Morosos > a 60 días"
         .Cells(r_int_NroFil + 15, 3) = "% Morosos > a 90 días"
         .Cells(r_int_NroFil + 16, 3) = "% Morosos > a 120 días"
      End If
      
      r_int_NumAux1 = 4
      r_int_NumAux2 = 1
      '12 MESES
      For r_int_TotFil = 1 To 12
          .Cells(r_int_NroFil + 1, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(1, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 2, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(2, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 3, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(3, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 4, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(4, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 5, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(5, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 6, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(6, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 7, r_int_NumAux1) = grd_LisRes.TextMatrix(7, r_int_NumAux2)
          .Cells(r_int_NroFil + 8, r_int_NumAux1) = grd_LisRes.TextMatrix(8, r_int_NumAux2)
          .Cells(r_int_NroFil + 9, r_int_NumAux1) = grd_LisRes.TextMatrix(9, r_int_NumAux2)
          .Cells(r_int_NroFil + 10, r_int_NumAux1) = grd_LisRes.TextMatrix(10, r_int_NumAux2)
          .Cells(r_int_NroFil + 11, r_int_NumAux1) = grd_LisRes.TextMatrix(11, r_int_NumAux2)
          .Cells(r_int_NroFil + 12, r_int_NumAux1) = grd_LisRes.TextMatrix(12, r_int_NumAux2)
          If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
             .Cells(r_int_NroFil + 13, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(13, r_int_NumAux2), "00.00")
             .Cells(r_int_NroFil + 14, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(14, r_int_NumAux2), "00.00")
             .Cells(r_int_NroFil + 15, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(15, r_int_NumAux2), "00.00")
             .Cells(r_int_NroFil + 16, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(16, r_int_NumAux2), "00.00")
          End If
          r_int_NumAux1 = r_int_NumAux1 + 1
          r_int_NumAux2 = r_int_NumAux2 + 1
      Next
 
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).Interior.Color = RGB(155, 187, 89)
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).HorizontalAlignment = xlHAlignCenter
            
      r_int_TotFil = 12
      If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
         r_int_TotFil = 16
      End If
      
      For r_int_Contador = 1 To r_int_TotFil
         If (r_int_Contador < 10) Then
             .Range(.Cells(r_int_Contador + 5, 3), .Cells(r_int_Contador + 5, 15)).NumberFormat = "##0.00"
         End If
         If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
            If (r_int_Contador > 12) Then
                .Range(.Cells(r_int_Contador + 5, 3), .Cells(r_int_Contador + 5, 15)).NumberFormat = "##0.00"
            End If
         End If
         If r_int_Contador Mod 2 <> 0 Then
            .Range("C" & r_int_Contador + 5 & ":O" & r_int_Contador + 5).Interior.Color = RGB(196, 215, 155)
         Else
            .Range("C" & r_int_Contador + 5 & ":O" & r_int_Contador + 5).Interior.Color = RGB(235, 241, 222)
         End If
      Next
         
      .Range("C5:O5").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C6:O6").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C7:O7").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C8:O8").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C9:O9").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C10:O10").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C11:O11").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C12:O12").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C13:O13").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C14:O14").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C15:O15").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C16:O16").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C17:O17").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C18:O18").Borders(xlEdgeTop).LineStyle = xlContinuous
      
      If moddat_g_int_TipRep <> 2 Then
         .Range("C19:O19").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C20:O20").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C21:O21").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C22:O22").Borders(xlEdgeTop).LineStyle = xlContinuous
      End If
      r_int_TotFil = 17
      If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
         r_int_TotFil = 21
      End If
      
      .Range("C5:C" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("D5:D" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("E5:E" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("F5:F" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("G5:G" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("H5:H" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("I5:I" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("J5:J" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("K5:K" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("L5:L" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("M5:M" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("N5:N" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("O5:O" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("P5:P" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
            
      .Range(.Columns("A"), .Columns("B")).ColumnWidth = 8
      .Range(.Columns("C"), .Columns("C")).ColumnWidth = 16
      .Range(.Columns("D"), .Columns("O")).ColumnWidth = 7
      .Columns("C").ColumnWidth = 17
      .Range("C15:O15").Font.Bold = True
      .Range("C15:O15").Interior.Color = RGB(155, 187, 89)
   End With
   '**************************************
   '**************************************
   'FORMATEA TITULO Y COLUMNAS GENERAL
   r_obj_Excel.Sheets(2).Name = "DETALLADO"
   With r_obj_Excel.Sheets(2)
      .Cells(3, 3) = "CARTERA ATRASADA: Vencido + Judicial/Total Crédito ó Tasa de Morosidad (%)"
      .Range(.Cells(3, 3), .Cells(3, 15)).Merge
      .Range("C3:O3").HorizontalAlignment = xlHAlignCenter
      .Range("C3:O3").Font.Bold = True
      .Range(.Cells(3, 3), .Cells(3, 15)).Font.Size = 14
      .Range("C4:O4").HorizontalAlignment = xlHAlignCenter
      .Range("C4:O4").Font.Bold = True
      .Cells(4, 3) = l_str_FecAno & "  -  " & Trim(l_str_NomPrd)
      .Range(.Cells(4, 3), .Cells(4, 15)).Merge
      .Range("C5:O25").Font.Name = "Arial Narrow"
      .Range("C5:O25").Font.Size = 10
      
      .Cells(r_int_NroFil, 3) = "PRODUCTO"
      .Cells(r_int_NroFil, 4) = "ENE"
      .Cells(r_int_NroFil, 5) = "FEB"
      .Cells(r_int_NroFil, 6) = "MAR"
      .Cells(r_int_NroFil, 7) = "ABR"
      .Cells(r_int_NroFil, 8) = "MAY"
      .Cells(r_int_NroFil, 9) = "JUN"
      .Cells(r_int_NroFil, 10) = "JUL"
      .Cells(r_int_NroFil, 11) = "AGO"
      .Cells(r_int_NroFil, 12) = "SET"
      .Cells(r_int_NroFil, 13) = "OCT"
      .Cells(r_int_NroFil, 14) = "NOV"
      .Cells(r_int_NroFil, 15) = "DIC"
      
      .Cells(r_int_NroFil + 1, 3) = "CRC-PBP"
      .Cells(r_int_NroFil + 2, 3) = "MICASITA"
      .Cells(r_int_NroFil + 3, 3) = "CME"
      .Cells(r_int_NroFil + 4, 3) = "MIVIVIENDA"
      .Cells(r_int_NroFil + 5, 3) = "MICASAMAS"
      .Cells(r_int_NroFil + 6, 3) = "BBP"
      .Cells(r_int_NroFil + 7, 3) = "TECHO PROPIO"
      .Cells(r_int_NroFil + 8, 3) = "TOTAL"
      .Cells(r_int_NroFil + 9, 3) = "CAR"
      .Cells(r_int_NroFil + 10, 3) = "TOTAL CANTIDAD"
      .Cells(r_int_NroFil + 11, 3) = "Morosos > a 90 días"
      .Cells(r_int_NroFil + 12, 3) = "Morosos > a 30 días"
      .Cells(r_int_NroFil + 13, 3) = "DESEMBOLSADOS"
      .Cells(r_int_NroFil + 14, 3) = "CANCELADOS"
      .Cells(r_int_NroFil + 15, 3) = "TRANSFERIDOS"
      '-------"REPORTE_INTERNO"--------"REPORTE_MICROEMPRESARIO"-------
      If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
         .Cells(r_int_NroFil + 16, 3) = "% Morosos > a 30 días"
         .Cells(r_int_NroFil + 17, 3) = "% Morosos > a 60 días"
         .Cells(r_int_NroFil + 18, 3) = "% Morosos > a 90 días"
         .Cells(r_int_NroFil + 19, 3) = "% Morosos > a 120 días"
      End If
      
      r_int_NumAux1 = 4
      r_int_NumAux2 = 1
      '12 MESES
      For r_int_TotFil = 1 To 12
          .Cells(r_int_NroFil + 1, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(1, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 2, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(2, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 3, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(3, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 4, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(4, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 5, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(5, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 6, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(6, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 7, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(7, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 8, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(8, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 9, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(9, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 10, r_int_NumAux1) = grd_LisDet.TextMatrix(10, r_int_NumAux2)
          .Cells(r_int_NroFil + 11, r_int_NumAux1) = grd_LisDet.TextMatrix(11, r_int_NumAux2)
          .Cells(r_int_NroFil + 12, r_int_NumAux1) = grd_LisDet.TextMatrix(12, r_int_NumAux2)
          .Cells(r_int_NroFil + 13, r_int_NumAux1) = grd_LisDet.TextMatrix(13, r_int_NumAux2)
          .Cells(r_int_NroFil + 14, r_int_NumAux1) = grd_LisDet.TextMatrix(14, r_int_NumAux2)
          .Cells(r_int_NroFil + 15, r_int_NumAux1) = grd_LisDet.TextMatrix(15, r_int_NumAux2)
          If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
             .Cells(r_int_NroFil + 16, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(16, r_int_NumAux2), "00.00")
             .Cells(r_int_NroFil + 17, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(17, r_int_NumAux2), "00.00")
             .Cells(r_int_NroFil + 18, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(18, r_int_NumAux2), "00.00")
             .Cells(r_int_NroFil + 19, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(19, r_int_NumAux2), "00.00")
          End If
          r_int_NumAux1 = r_int_NumAux1 + 1
          r_int_NumAux2 = r_int_NumAux2 + 1
      Next

      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).Interior.Color = RGB(155, 187, 89)
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).HorizontalAlignment = xlHAlignCenter
            
      r_int_TotFil = 15
      If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
         r_int_TotFil = 19
      End If
      
      For r_int_Contador = 1 To r_int_TotFil
         If (r_int_Contador < 10) Then
             .Range(.Cells(r_int_Contador + 5, 3), .Cells(r_int_Contador + 5, 15)).NumberFormat = "##0.00"
         End If
         If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
            If (r_int_Contador > 15) Then
                .Range(.Cells(r_int_Contador + 5, 3), .Cells(r_int_Contador + 5, 15)).NumberFormat = "##0.00"
            End If
         End If
         If r_int_Contador Mod 2 <> 0 Then
            .Range("C" & r_int_Contador + 5 & ":O" & r_int_Contador + 5).Interior.Color = RGB(196, 215, 155)
         Else
            .Range("C" & r_int_Contador + 5 & ":O" & r_int_Contador + 5).Interior.Color = RGB(235, 241, 222)
         End If
      Next
         
      .Range("C5:O5").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C6:O6").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C7:O7").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C8:O8").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C9:O9").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C10:O10").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C11:O11").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C12:O12").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C13:O13").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C14:O14").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C15:O15").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C16:O16").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C17:O17").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C18:O18").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C19:O19").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C20:O20").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C21:O21").Borders(xlEdgeTop).LineStyle = xlContinuous
      
      If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
         .Range("C22:O22").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C23:O23").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C24:O24").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C25:O25").Borders(xlEdgeTop).LineStyle = xlContinuous
      End If
      
      r_int_TotFil = 20
      If moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5 Then
         r_int_TotFil = 24
      End If
      
      .Range("C5:C" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("D5:D" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("E5:E" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("F5:F" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("G5:G" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("H5:H" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("I5:I" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("J5:J" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("K5:K" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("L5:L" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("M5:M" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("N5:N" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("O5:O" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("P5:P" & r_int_TotFil).Borders(xlEdgeLeft).LineStyle = xlContinuous
            
      .Range(.Columns("A"), .Columns("B")).ColumnWidth = 8
      .Range(.Columns("C"), .Columns("C")).ColumnWidth = 16
      .Range(.Columns("D"), .Columns("O")).ColumnWidth = 7
      .Columns("C").ColumnWidth = 17
      .Range("C15:O15").Font.Bold = True
      .Range("C15:O15").Interior.Color = RGB(155, 187, 89)
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc2()
Dim r_obj_Excel     As Excel.Application
Dim r_int_NroFil    As Integer
Dim r_int_Contador  As Integer
Dim r_int_NumAux1   As Integer
Dim r_int_NumAux2   As Integer
Dim r_int_TotFil    As Integer
      
   'Preparando Cabecera de Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 5
         
   '**************************************
   'FORMATEA TITULO Y COLUMNAS GENERAL
   r_obj_Excel.Sheets(2).Name = "DETALLADO"
   With r_obj_Excel.Sheets(2)
      .Cells(3, 3) = "CARTERA ATRASADA: Vencido + Judicial/Total Crédito ó Tasa de Morosidad (%)"
      .Range("C3:O3").HorizontalAlignment = xlHAlignCenter
      .Range("C3:O3").Font.Bold = True
      .Range(.Cells(3, 3), .Cells(3, 15)).Font.Size = 14
      .Range(.Cells(3, 3), .Cells(3, 15)).Merge
            
      .Range("C4:O4").HorizontalAlignment = xlHAlignCenter
      .Range("C4:O4").Font.Bold = True
      .Cells(4, 3) = l_str_FecAno & "  -  " & Trim(l_str_NomPrd)
      .Range(.Cells(4, 3), .Cells(4, 15)).Merge
      
      .Range("C5:O15").Font.Name = "Arial Narrow"
      .Range("C5:O15").Font.Size = 10
      
      .Cells(r_int_NroFil, 3) = "PRODUCTO"
      .Cells(r_int_NroFil, 4) = "ENE"
      .Cells(r_int_NroFil, 5) = "FEB"
      .Cells(r_int_NroFil, 6) = "MAR"
      .Cells(r_int_NroFil, 7) = "ABR"
      .Cells(r_int_NroFil, 8) = "MAY"
      .Cells(r_int_NroFil, 9) = "JUN"
      .Cells(r_int_NroFil, 10) = "JUL"
      .Cells(r_int_NroFil, 11) = "AGO"
      .Cells(r_int_NroFil, 12) = "SET"
      .Cells(r_int_NroFil, 13) = "OCT"
      .Cells(r_int_NroFil, 14) = "NOV"
      .Cells(r_int_NroFil, 15) = "DIC"
      
      .Cells(r_int_NroFil + 1, 3) = grd_LisDet.TextMatrix(1, 0)
      .Cells(r_int_NroFil + 2, 3) = grd_LisDet.TextMatrix(2, 0)
      .Cells(r_int_NroFil + 3, 3) = grd_LisDet.TextMatrix(3, 0)
      .Cells(r_int_NroFil + 4, 3) = grd_LisDet.TextMatrix(4, 0)
      .Cells(r_int_NroFil + 5, 3) = "TOTAL CANTIDAD"
      .Cells(r_int_NroFil + 6, 3) = "Morosos > a 90 días"
      .Cells(r_int_NroFil + 7, 3) = "Morosos > a 0 días"
      .Cells(r_int_NroFil + 8, 3) = "DESEMBOLSADOS"
      .Cells(r_int_NroFil + 9, 3) = "CANCELADOS"
      .Cells(r_int_NroFil + 10, 3) = "TRANSFERIDOS"
      
      r_int_NumAux1 = 4
      r_int_NumAux2 = 1
      '12 MESES
      For r_int_TotFil = 1 To 12
          .Cells(r_int_NroFil + 1, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(1, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 2, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(2, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 3, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(3, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 4, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(4, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 5, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(5, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 6, r_int_NumAux1) = Format(grd_LisDet.TextMatrix(6, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 7, r_int_NumAux1) = grd_LisDet.TextMatrix(7, r_int_NumAux2)
          .Cells(r_int_NroFil + 8, r_int_NumAux1) = grd_LisDet.TextMatrix(8, r_int_NumAux2)
          .Cells(r_int_NroFil + 9, r_int_NumAux1) = grd_LisDet.TextMatrix(9, r_int_NumAux2)
          .Cells(r_int_NroFil + 10, r_int_NumAux1) = grd_LisDet.TextMatrix(10, r_int_NumAux2)
          r_int_NumAux1 = r_int_NumAux1 + 1
          r_int_NumAux2 = r_int_NumAux2 + 1
      Next
      
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).Interior.Color = RGB(155, 187, 89)
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).HorizontalAlignment = xlHAlignCenter
      
      For r_int_Contador = 1 To 10
         If (r_int_Contador < 5) Then
             .Range(.Cells(r_int_Contador + 5, 3), .Cells(r_int_Contador + 5, 15)).NumberFormat = "##0.00"
         End If
         If r_int_Contador Mod 2 <> 0 Then
            .Range("C" & r_int_Contador + 5 & ":O" & r_int_Contador + 5).Interior.Color = RGB(196, 215, 155)
         Else
            .Range("C" & r_int_Contador + 5 & ":O" & r_int_Contador + 5).Interior.Color = RGB(235, 241, 222)
         End If
      Next
         
      .Range("C5:O5").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C6:O6").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C7:O7").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C8:O8").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C9:O9").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C10:O10").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C11:O11").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C12:O12").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C13:O13").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C14:O14").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C15:O15").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C16:O16").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C5:C15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("D5:D15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("E5:E15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("F5:F15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("G5:G15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("H5:H15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("I5:I15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("J5:J15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("K5:K15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("L5:L15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("M5:M15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("N5:N15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("O5:O15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("P5:P15").Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      .Range(.Columns("A"), .Columns("B")).ColumnWidth = 8
      .Range(.Columns("C"), .Columns("C")).ColumnWidth = 16
      .Range(.Columns("D"), .Columns("O")).ColumnWidth = 7
      .Range("C10:O10").Font.Bold = True
      .Range("C10:O10").Interior.Color = RGB(155, 187, 89)
   End With
   
   '**************************************
   'FORMATEA TITULO Y COLUMNAS GENERAL
   r_obj_Excel.Sheets(1).Name = "RESUMEN"
   With r_obj_Excel.Sheets(1)
      .Cells(3, 3) = "CARTERA ATRASADA: Vencido + Judicial/Total Crédito ó Tasa de Morosidad (%)"
      .Range("C3:O3").Select
      .Range("C3:O3").HorizontalAlignment = xlHAlignCenter
      .Range("C3:O3").Font.Bold = True
      .Range(.Cells(3, 3), .Cells(3, 15)).Font.Size = 14
      r_obj_Excel.Selection.MergeCells = True
      
      .Range("C4:O4").Select
      .Range("C4:O4").HorizontalAlignment = xlHAlignCenter
      .Range("C4:O4").Font.Bold = True
   
      .Cells(4, 15) = l_str_FecAno & "  -  " & Trim(l_str_NomPrd)
      r_obj_Excel.Selection.MergeCells = True
      
      .Range("C5:O15").Font.Name = "Arial Narrow"
      .Range("C5:O15").Font.Size = 10
      
      .Cells(r_int_NroFil, 3) = "PRODUCTO"
      .Cells(r_int_NroFil, 4) = "ENE"
      .Cells(r_int_NroFil, 5) = "FEB"
      .Cells(r_int_NroFil, 6) = "MAR"
      .Cells(r_int_NroFil, 7) = "ABR"
      .Cells(r_int_NroFil, 8) = "MAY"
      .Cells(r_int_NroFil, 9) = "JUN"
      .Cells(r_int_NroFil, 10) = "JUL"
      .Cells(r_int_NroFil, 11) = "AGO"
      .Cells(r_int_NroFil, 12) = "SET"
      .Cells(r_int_NroFil, 13) = "OCT"
      .Cells(r_int_NroFil, 14) = "NOV"
      .Cells(r_int_NroFil, 15) = "DIC"
      
      .Cells(r_int_NroFil + 1, 3) = grd_LisRes.TextMatrix(1, 0)
      .Cells(r_int_NroFil + 2, 3) = grd_LisRes.TextMatrix(2, 0)
      .Cells(r_int_NroFil + 3, 3) = grd_LisRes.TextMatrix(3, 0)
      .Cells(r_int_NroFil + 4, 3) = "TOTAL CANTIDAD"
      .Cells(r_int_NroFil + 5, 3) = "Morosos > a 90 días"
      .Cells(r_int_NroFil + 6, 3) = "Morosos > a 0 días"
      .Cells(r_int_NroFil + 7, 3) = "DESEMBOLSADOS"
      .Cells(r_int_NroFil + 8, 3) = "CANCELADOS"
      .Cells(r_int_NroFil + 9, 3) = "TRANSFERIDOS"
      
      r_int_NumAux1 = 4
      r_int_NumAux2 = 1
      '12 MESES
      For r_int_TotFil = 1 To 12
          .Cells(r_int_NroFil + 1, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(1, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 2, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(2, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 3, r_int_NumAux1) = Format(grd_LisRes.TextMatrix(3, r_int_NumAux2), "00.00")
          .Cells(r_int_NroFil + 4, r_int_NumAux1) = grd_LisRes.TextMatrix(4, r_int_NumAux2)
          .Cells(r_int_NroFil + 5, r_int_NumAux1) = grd_LisRes.TextMatrix(5, r_int_NumAux2)
          .Cells(r_int_NroFil + 6, r_int_NumAux1) = grd_LisRes.TextMatrix(6, r_int_NumAux2)
          .Cells(r_int_NroFil + 7, r_int_NumAux1) = grd_LisRes.TextMatrix(7, r_int_NumAux2)
          .Cells(r_int_NroFil + 8, r_int_NumAux1) = grd_LisRes.TextMatrix(8, r_int_NumAux2)
          .Cells(r_int_NroFil + 9, r_int_NumAux1) = grd_LisRes.TextMatrix(9, r_int_NumAux2)
          r_int_NumAux1 = r_int_NumAux1 + 1
          r_int_NumAux2 = r_int_NumAux2 + 1
      Next
      
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).Interior.Color = RGB(155, 187, 89)
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 15)).HorizontalAlignment = xlHAlignCenter
      
      For r_int_Contador = 1 To 9
         If (r_int_Contador < 4) Then
             .Range(.Cells(r_int_Contador + 5, 3), .Cells(r_int_Contador + 5, 15)).NumberFormat = "##0.00"
         End If
         If r_int_Contador Mod 2 <> 0 Then
            .Range("C" & r_int_Contador + 5 & ":O" & r_int_Contador + 5).Interior.Color = RGB(196, 215, 155)
         Else
            .Range("C" & r_int_Contador + 5 & ":O" & r_int_Contador + 5).Interior.Color = RGB(235, 241, 222)
         End If
      Next
         
      .Range("C5:O5").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C6:O6").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C7:O7").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C8:O8").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C9:O9").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C10:O10").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C11:O11").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C12:O12").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C13:O13").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C14:O14").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C15:O15").Borders(xlEdgeTop).LineStyle = xlContinuous
      
      .Range("C5:C14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("D5:D14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("E5:E14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("F5:F14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("G5:G14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("H5:H14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("I5:I14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("J5:J14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("K5:K14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("L5:L14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("M5:M14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("N5:N14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("O5:O14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range("P5:P14").Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      .Range(.Columns("A"), .Columns("B")).ColumnWidth = 8
      .Range(.Columns("C"), .Columns("C")).ColumnWidth = 16
      .Range(.Columns("D"), .Columns("O")).ColumnWidth = 7
      .Range("C10:O10").Font.Bold = True
      .Range("C10:O10").Interior.Color = RGB(155, 187, 89)
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_LisDet_DblClick()
   If grd_LisDet.Rows = 2 Then
      Exit Sub
   End If
   
   moddat_g_str_CodAno = CInt(ipp_PerAno.Text)
   moddat_g_str_FecIni = l_str_FecIni
   moddat_g_str_FecFin = l_str_FecCal
   moddat_g_str_FecCan = l_str_FecLim
   moddat_g_str_CodMes = l_str_PerMes
   
   'CRC_PBP = 1, MICASITA = 2, CME = 3, MIVIVIENDA = 4, MICASAMAS = 5, BBP = 6 , TECHO PROPIO = 7 , MAYOR_90 = 8, MAYOR_0 = 9
   If (moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 2 Or moddat_g_int_TipRep = 5) Then
       If (grd_LisDet.RowSel = 8 Or grd_LisDet.RowSel = 9 Or grd_LisDet.RowSel = 10) Then
           Exit Sub
       End If
       moddat_g_str_NomPrd = UCase(Trim(grd_LisDet.TextMatrix(grd_LisDet.RowSel, 0)) & " -  período " & (l_str_FecAno))
       
       Select Case grd_LisDet.RowSel
         Case 1: moddat_g_str_TipPar = 1 'CRC_PBP
         Case 2: moddat_g_str_TipPar = 2 'MICASITA
         Case 3: moddat_g_str_TipPar = 3 'CME
         Case 4: moddat_g_str_TipPar = 4 'MIVIVIENDA
         
         Case 5: moddat_g_str_TipPar = 5 'MICASAMAS
         Case 6: moddat_g_str_TipPar = 6 'BBP
         Case 7: moddat_g_str_TipPar = 7 'TECHO PROPIO
         
         Case 11: moddat_g_str_TipPar = 8 'MAYOR_90
         Case 12: moddat_g_str_TipPar = 9 'MAYOR_0
         Case 13, 14, 15: moddat_g_str_TipPar = 10
       End Select
   Else
       If (grd_LisDet.RowSel = 3 Or grd_LisDet.RowSel = 4 Or grd_LisDet.RowSel = 5) Then
           Exit Sub
       End If
       moddat_g_str_NomPrd = UCase(Trim(grd_LisDet.TextMatrix(grd_LisDet.RowSel, 0)) & "-  período " & (l_str_FecAno))
       If (moddat_g_int_TipRep = 3) Then 'MIVIVIENDA
           Select Case grd_LisDet.RowSel
                  Case 1: moddat_g_str_TipPar = 1 'MIVIVIENDA
                  Case 2: moddat_g_str_TipPar = 3 'CME
                  Case 6: moddat_g_str_TipPar = 5 'MAYOR_90
                  Case 7: moddat_g_str_TipPar = 6 'MAYOR_0
                  Case 8, 9, 10: moddat_g_str_TipPar = 7
           End Select
       ElseIf (moddat_g_int_TipRep = 4) Then 'MICASITA
           Select Case grd_LisDet.RowSel
                  Case 1: moddat_g_str_TipPar = 2 'MICASITA
                  Case 2: moddat_g_str_TipPar = 4 'MICASITA_PBP
                  Case 6: moddat_g_str_TipPar = 5 'MAYOR_90
                  Case 7: moddat_g_str_TipPar = 6 'MAYOR_0
                  Case 8, 9, 10: moddat_g_str_TipPar = 7
           End Select
       End If
   End If
   
   moddat_g_int_OrdAct = 0
   If (moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 5) Then
      If grd_LisDet.RowSel = 16 Or grd_LisDet.RowSel = 17 Or grd_LisDet.RowSel = 18 Or grd_LisDet.RowSel = 19 Then
         Select Case grd_LisDet.RowSel
            Case 16: moddat_g_str_TipPar = 11 '> 30 DIAS
            Case 17: moddat_g_str_TipPar = 12 '> 60 DIAS
            Case 18: moddat_g_str_TipPar = 13 '> 90 DIAS
            Case 19: moddat_g_str_TipPar = 14 '> 120 DIAS
         End Select
      Else
         If (moddat_g_str_TipPar = 8) Then
             moddat_g_int_OrdAct = 90
         End If
         If (moddat_g_str_TipPar = 9) Then
             moddat_g_int_OrdAct = 30
         End If
      End If
   ElseIf (moddat_g_int_TipRep = 2) Then
       If (moddat_g_str_TipPar = 5) Then
           moddat_g_int_OrdAct = 90
       End If
       If (moddat_g_str_TipPar = 6) Then
           moddat_g_int_OrdAct = 0
       End If
   Else
       If (moddat_g_str_TipPar = 5) Then
           moddat_g_int_OrdAct = 90
       End If
       If (moddat_g_str_TipPar = 6) Then
           moddat_g_int_OrdAct = 0
       End If
       
       
   End If
   
   If grd_LisDet.RowSel <> 8 Then
      frm_RptCtb_31.Show 1
   End If
End Sub

Private Sub grd_LisDet_SelChange()
   If grd_LisDet.Rows > 2 Then
      grd_LisDet.RowSel = grd_LisDet.Row
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If CInt(ipp_PerAno.Text) >= 2010 Then
         Call gs_SetFocus(cmb_TipDes)
      End If
   End If
End Sub

Private Sub cmb_TipDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipDes.ListIndex > -1 Then
         Call gs_SetFocus(cmd_Proces)
      End If
   End If
End Sub

