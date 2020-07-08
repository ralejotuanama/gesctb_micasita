VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_24 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   1425
   ClientTop       =   2175
   ClientWidth     =   14085
   Icon            =   "GesCtb_frm_850.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   14085
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel5 
      Height          =   8775
      Left            =   -90
      TabIndex        =   6
      Top             =   0
      Width           =   14175
      _Version        =   65536
      _ExtentX        =   25003
      _ExtentY        =   15478
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   150
         TabIndex        =   7
         Top             =   60
         Width           =   13965
         _Version        =   65536
         _ExtentX        =   24642
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   300
            Left            =   600
            TabIndex        =   8
            Top             =   180
            Width           =   7335
            _Version        =   65536
            _ExtentX        =   12938
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte de Consolidado de Clasificaciones de Cartera para Provisiones"
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
            Picture         =   "GesCtb_frm_850.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   150
         TabIndex        =   9
         Top             =   765
         Width           =   13965
         _Version        =   65536
         _ExtentX        =   24633
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
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   45
            Picture         =   "GesCtb_frm_850.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Procesar informacion"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExcRes 
            Height          =   585
            Left            =   645
            Picture         =   "GesCtb_frm_850.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13320
            Picture         =   "GesCtb_frm_850.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   45
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   825
         Left            =   180
         TabIndex        =   10
         Top             =   1470
         Width           =   13965
         _Version        =   65536
         _ExtentX        =   24633
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
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   3795
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   6780
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   90
            Width           =   3795
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1200
            TabIndex        =   2
            Top             =   420
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
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   5700
            TabIndex        =   12
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   450
            Width           =   975
         End
      End
      Begin TabDlg.SSTab tab_Clasif 
         Height          =   6315
         Left            =   180
         TabIndex        =   14
         Top             =   2370
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   11139
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Resumen"
         TabPicture(0)   =   "GesCtb_frm_850.frx":0D6C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "grd_LisCab"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detallado"
         TabPicture(1)   =   "GesCtb_frm_850.frx":0D88
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "grd_LisDet"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSFlexGridLib.MSFlexGrid grd_LisDet 
            Height          =   5775
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   13680
            _ExtentX        =   24130
            _ExtentY        =   10186
            _Version        =   393216
            Rows            =   21
            Cols            =   38
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            Redraw          =   -1  'True
            MergeCells      =   1
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisCab 
            Height          =   5775
            Left            =   -74880
            TabIndex        =   16
            Top             =   360
            Width           =   13680
            _ExtentX        =   24130
            _ExtentY        =   10186
            _Version        =   393216
            Rows            =   12
            Cols            =   38
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            Redraw          =   -1  'True
            MergeCells      =   1
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
End
Attribute VB_Name = "frm_RptCtb_24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_int_Produc     As Integer
Dim r_int_PerMes     As Integer
Dim r_int_PerAno     As Integer

Private Sub cmd_ExpExcRes_Click()
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
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExcRes
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Proces_Click()
   'Validaciones
   If cmb_Produc.ListIndex = -1 Then
      MsgBox "Debe seleccionar un producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Produc)
      Exit Sub
   End If
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   r_int_Produc = cmb_Produc.ItemData(cmb_Produc.ListIndex)
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   
   Call fs_Obtiene_Cabecera
   Call fs_Obtiene_Detalle
      
   Call fs_Activa(True)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(False)
   Call gs_CentraForm(Me)
   
   tab_Clasif.Tab = 0
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
   
   cmb_Produc.Clear
   cmb_Produc.AddItem "- TODOS -"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 0
   cmb_Produc.AddItem "CRC-PBP"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 1
   cmb_Produc.AddItem "MICASITA"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 2
   cmb_Produc.AddItem "CME"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 3
   cmb_Produc.AddItem "MIVIVIENDA"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 4
   
   cmb_Produc.AddItem "MICASAMAS"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 5
   cmb_Produc.AddItem "BBP"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 6
   cmb_Produc.AddItem "TECHO PROPIO"
   cmb_Produc.ItemData(cmb_Produc.NewIndex) = 7
   cmb_Produc.ListIndex = 0
   
   grd_LisCab.ColWidth(0) = 0       ' CODIGO DE CLASIFICACION
   grd_LisCab.ColWidth(1) = 2000    ' DESCRIPCION DE CLASIFICACION
   grd_LisCab.ColWidth(2) = 900     ' NUMERO MES 01
   grd_LisCab.ColWidth(3) = 1300    ' MONTO MES 01
   grd_LisCab.ColWidth(4) = 1300    ' MONTO MES 01
   grd_LisCab.ColWidth(5) = 900     ' NUMERO MES 02
   grd_LisCab.ColWidth(6) = 1300    ' MONTO MES 02
   grd_LisCab.ColWidth(7) = 1300    ' MONTO MES 02
   grd_LisCab.ColWidth(8) = 900     ' NUMERO MES 03
   grd_LisCab.ColWidth(9) = 1300    ' MONTO MES 03
   grd_LisCab.ColWidth(10) = 1300    ' MONTO MES 03
   grd_LisCab.ColWidth(11) = 900     ' NUMERO MES 04
   grd_LisCab.ColWidth(12) = 1300    ' MONTO MES 04
   grd_LisCab.ColWidth(13) = 1300    ' MONTO MES 04
   grd_LisCab.ColWidth(14) = 900     ' NUMERO MES 05
   grd_LisCab.ColWidth(15) = 1300    ' MONTO MES 05
   grd_LisCab.ColWidth(16) = 1300    ' MONTO MES 05
   grd_LisCab.ColWidth(17) = 900     ' NUMERO MES 06
   grd_LisCab.ColWidth(18) = 1300    ' MONTO MES 06
   grd_LisCab.ColWidth(19) = 1300    ' MONTO MES 06
   grd_LisCab.ColWidth(20) = 900     ' NUMERO MES 07
   grd_LisCab.ColWidth(21) = 1300    ' MONTO MES 07
   grd_LisCab.ColWidth(22) = 1300    ' MONTO MES 07
   grd_LisCab.ColWidth(23) = 900     ' NUMERO MES 08
   grd_LisCab.ColWidth(24) = 1300    ' MONTO MES 08
   grd_LisCab.ColWidth(25) = 1300    ' MONTO MES 08
   grd_LisCab.ColWidth(26) = 900     ' NUMERO MES 09
   grd_LisCab.ColWidth(27) = 1300    ' MONTO MES 09
   grd_LisCab.ColWidth(28) = 1300    ' MONTO MES 09
   grd_LisCab.ColWidth(29) = 900     ' NUMERO MES 10
   grd_LisCab.ColWidth(30) = 1300    ' MONTO MES 10
   grd_LisCab.ColWidth(31) = 1300    ' MONTO MES 10
   grd_LisCab.ColWidth(32) = 900     ' NUMERO MES 11
   grd_LisCab.ColWidth(33) = 1300    ' MONTO MES 11
   grd_LisCab.ColWidth(34) = 1300    ' MONTO MES 11
   grd_LisCab.ColWidth(35) = 900     ' NUMERO MES 12
   grd_LisCab.ColWidth(36) = 1300    ' MONTO MES 12
   grd_LisCab.ColWidth(37) = 1300    ' MONTO MES 12
   grd_LisCab.ColAlignment(0) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(1) = flexAlignLeftCenter
   grd_LisCab.ColAlignment(2) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(3) = flexAlignRightCenter
   grd_LisCab.ColAlignment(4) = flexAlignRightCenter
   grd_LisCab.ColAlignment(5) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(6) = flexAlignRightCenter
   grd_LisCab.ColAlignment(7) = flexAlignRightCenter
   grd_LisCab.ColAlignment(8) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(9) = flexAlignRightCenter
   grd_LisCab.ColAlignment(10) = flexAlignRightCenter
   grd_LisCab.ColAlignment(11) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(12) = flexAlignRightCenter
   grd_LisCab.ColAlignment(13) = flexAlignRightCenter
   grd_LisCab.ColAlignment(14) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(15) = flexAlignRightCenter
   grd_LisCab.ColAlignment(16) = flexAlignRightCenter
   grd_LisCab.ColAlignment(17) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(18) = flexAlignRightCenter
   grd_LisCab.ColAlignment(19) = flexAlignRightCenter
   grd_LisCab.ColAlignment(20) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(21) = flexAlignRightCenter
   grd_LisCab.ColAlignment(22) = flexAlignRightCenter
   grd_LisCab.ColAlignment(23) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(24) = flexAlignRightCenter
   grd_LisCab.ColAlignment(25) = flexAlignRightCenter
   grd_LisCab.ColAlignment(26) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(27) = flexAlignRightCenter
   grd_LisCab.ColAlignment(28) = flexAlignRightCenter
   grd_LisCab.ColAlignment(29) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(30) = flexAlignRightCenter
   grd_LisCab.ColAlignment(31) = flexAlignRightCenter
   grd_LisCab.ColAlignment(32) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(33) = flexAlignRightCenter
   grd_LisCab.ColAlignment(34) = flexAlignRightCenter
   grd_LisCab.ColAlignment(35) = flexAlignCenterCenter
   grd_LisCab.ColAlignment(36) = flexAlignRightCenter
   grd_LisCab.ColAlignment(37) = flexAlignRightCenter
   Call gs_LimpiaGrid(grd_LisCab)
   
   grd_LisDet.ColWidth(0) = 0       ' CODIGO DE CLASIFICACION
   grd_LisDet.ColWidth(1) = 2000    ' DESCRIPCION DE CLASIFICACION
   grd_LisDet.ColWidth(2) = 900     ' NUMERO MES 01
   grd_LisDet.ColWidth(3) = 1300    ' MONTO MES 01
   grd_LisDet.ColWidth(4) = 1300    ' MONTO MES 01
   grd_LisDet.ColWidth(5) = 900     ' NUMERO MES 02
   grd_LisDet.ColWidth(6) = 1300    ' MONTO MES 02
   grd_LisDet.ColWidth(7) = 1300    ' MONTO MES 02
   grd_LisDet.ColWidth(8) = 900     ' NUMERO MES 03
   grd_LisDet.ColWidth(9) = 1300    ' MONTO MES 03
   grd_LisDet.ColWidth(10) = 1300    ' MONTO MES 03
   grd_LisDet.ColWidth(11) = 900     ' NUMERO MES 04
   grd_LisDet.ColWidth(12) = 1300    ' MONTO MES 04
   grd_LisDet.ColWidth(13) = 1300    ' MONTO MES 04
   grd_LisDet.ColWidth(14) = 900     ' NUMERO MES 05
   grd_LisDet.ColWidth(15) = 1300    ' MONTO MES 05
   grd_LisDet.ColWidth(16) = 1300    ' MONTO MES 05
   grd_LisDet.ColWidth(17) = 900     ' NUMERO MES 06
   grd_LisDet.ColWidth(18) = 1300    ' MONTO MES 06
   grd_LisDet.ColWidth(19) = 1300    ' MONTO MES 06
   grd_LisDet.ColWidth(20) = 900     ' NUMERO MES 07
   grd_LisDet.ColWidth(21) = 1300    ' MONTO MES 07
   grd_LisDet.ColWidth(22) = 1300    ' MONTO MES 07
   grd_LisDet.ColWidth(23) = 900     ' NUMERO MES 08
   grd_LisDet.ColWidth(24) = 1300    ' MONTO MES 08
   grd_LisDet.ColWidth(25) = 1300    ' MONTO MES 08
   grd_LisDet.ColWidth(26) = 900     ' NUMERO MES 09
   grd_LisDet.ColWidth(27) = 1300    ' MONTO MES 09
   grd_LisDet.ColWidth(28) = 1300    ' MONTO MES 09
   grd_LisDet.ColWidth(29) = 900     ' NUMERO MES 10
   grd_LisDet.ColWidth(30) = 1300    ' MONTO MES 10
   grd_LisDet.ColWidth(31) = 1300    ' MONTO MES 10
   grd_LisDet.ColWidth(32) = 900     ' NUMERO MES 11
   grd_LisDet.ColWidth(33) = 1300    ' MONTO MES 11
   grd_LisDet.ColWidth(34) = 1300    ' MONTO MES 11
   grd_LisDet.ColWidth(35) = 900     ' NUMERO MES 12
   grd_LisDet.ColWidth(36) = 1300    ' MONTO MES 12
   grd_LisDet.ColWidth(37) = 1300    ' MONTO MES 12
   grd_LisDet.ColAlignment(0) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(1) = flexAlignLeftCenter
   grd_LisDet.ColAlignment(2) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(3) = flexAlignRightCenter
   grd_LisDet.ColAlignment(4) = flexAlignRightCenter
   grd_LisDet.ColAlignment(5) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(6) = flexAlignRightCenter
   grd_LisDet.ColAlignment(7) = flexAlignRightCenter
   grd_LisDet.ColAlignment(8) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(9) = flexAlignRightCenter
   grd_LisDet.ColAlignment(10) = flexAlignRightCenter
   grd_LisDet.ColAlignment(11) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(12) = flexAlignRightCenter
   grd_LisDet.ColAlignment(13) = flexAlignRightCenter
   grd_LisDet.ColAlignment(14) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(15) = flexAlignRightCenter
   grd_LisDet.ColAlignment(16) = flexAlignRightCenter
   grd_LisDet.ColAlignment(17) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(18) = flexAlignRightCenter
   grd_LisDet.ColAlignment(19) = flexAlignRightCenter
   grd_LisDet.ColAlignment(20) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(21) = flexAlignRightCenter
   grd_LisDet.ColAlignment(22) = flexAlignRightCenter
   grd_LisDet.ColAlignment(23) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(24) = flexAlignRightCenter
   grd_LisDet.ColAlignment(25) = flexAlignRightCenter
   grd_LisDet.ColAlignment(26) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(27) = flexAlignRightCenter
   grd_LisDet.ColAlignment(28) = flexAlignRightCenter
   grd_LisDet.ColAlignment(29) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(30) = flexAlignRightCenter
   grd_LisDet.ColAlignment(31) = flexAlignRightCenter
   grd_LisDet.ColAlignment(32) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(33) = flexAlignRightCenter
   grd_LisDet.ColAlignment(34) = flexAlignRightCenter
   grd_LisDet.ColAlignment(35) = flexAlignCenterCenter
   grd_LisDet.ColAlignment(36) = flexAlignRightCenter
   grd_LisDet.ColAlignment(37) = flexAlignRightCenter
   Call gs_LimpiaGrid(grd_LisDet)
End Sub

Private Sub fs_Activa(ByVal estado As Boolean)
    cmd_ExpExcRes.Enabled = estado
End Sub

Private Sub fs_Obtiene_Cabecera()
Dim r_int_Contad     As Integer
Dim r_int_ConCla     As Integer
Dim r_int_NumCol     As Integer
Dim r_int_NumFil     As Integer
Dim r_int_TotNum     As Integer
Dim r_dbl_TotMto     As Double
Dim r_dbl_TotPrv     As Double
Dim r_int_Numero     As Integer
Dim r_dbl_MtoTot     As Double
Dim r_dbl_MtoPrv     As Double
Dim r_int_NumAli     As Integer
Dim r_dbl_TotAli     As Double
Dim r_dbl_PrvAli     As Double

   grd_LisCab.Redraw = False
   Call gs_LimpiaGrid(grd_LisCab)
   
   'Primera Linea
   grd_LisCab.Rows = grd_LisCab.Rows + 1
   grd_LisCab.Row = grd_LisCab.Rows - 1
   grd_LisCab.Row = 0:   grd_LisCab.Text = ""
   grd_LisCab.Col = 1:   grd_LisCab.Text = "CLASIFICACIONES"
   grd_LisCab.Col = 2:   grd_LisCab.Text = "ENERO"
   grd_LisCab.Col = 3:   grd_LisCab.Text = "ENERO"
   grd_LisCab.Col = 4:   grd_LisCab.Text = "ENERO"
   grd_LisCab.Col = 5:   grd_LisCab.Text = "FEBRERO"
   grd_LisCab.Col = 6:   grd_LisCab.Text = "FEBRERO"
   grd_LisCab.Col = 7:   grd_LisCab.Text = "FEBRERO"
   grd_LisCab.Col = 8:   grd_LisCab.Text = "MARZO"
   grd_LisCab.Col = 9:   grd_LisCab.Text = "MARZO"
   grd_LisCab.Col = 10:  grd_LisCab.Text = "MARZO"
   grd_LisCab.Col = 11:  grd_LisCab.Text = "ABRIL"
   grd_LisCab.Col = 12:  grd_LisCab.Text = "ABRIL"
   grd_LisCab.Col = 13:  grd_LisCab.Text = "ABRIL"
   grd_LisCab.Col = 14:  grd_LisCab.Text = "MAYO"
   grd_LisCab.Col = 15:  grd_LisCab.Text = "MAYO"
   grd_LisCab.Col = 16:  grd_LisCab.Text = "MAYO"
   grd_LisCab.Col = 17:  grd_LisCab.Text = "JUNIO"
   grd_LisCab.Col = 18:  grd_LisCab.Text = "JUNIO"
   grd_LisCab.Col = 19:  grd_LisCab.Text = "JUNIO"
   grd_LisCab.Col = 20:  grd_LisCab.Text = "JULIO"
   grd_LisCab.Col = 21:  grd_LisCab.Text = "JULIO"
   grd_LisCab.Col = 22:  grd_LisCab.Text = "JULIO"
   grd_LisCab.Col = 23:  grd_LisCab.Text = "AGOSTO"
   grd_LisCab.Col = 24:  grd_LisCab.Text = "AGOSTO"
   grd_LisCab.Col = 25:  grd_LisCab.Text = "AGOSTO"
   grd_LisCab.Col = 26:  grd_LisCab.Text = "SETIEMBRE"
   grd_LisCab.Col = 27:  grd_LisCab.Text = "SETIEMBRE"
   grd_LisCab.Col = 28:  grd_LisCab.Text = "SETIEMBRE"
   grd_LisCab.Col = 29:  grd_LisCab.Text = "OCTUBRE"
   grd_LisCab.Col = 30:  grd_LisCab.Text = "OCTUBRE"
   grd_LisCab.Col = 31:  grd_LisCab.Text = "OCTUBRE"
   grd_LisCab.Col = 32:  grd_LisCab.Text = "NOVIEMBRE"
   grd_LisCab.Col = 33:  grd_LisCab.Text = "NOVIEMBRE"
   grd_LisCab.Col = 34:  grd_LisCab.Text = "NOVIEMBRE"
   grd_LisCab.Col = 35:  grd_LisCab.Text = "DICIEMBRE"
   grd_LisCab.Col = 36:  grd_LisCab.Text = "DICIEMBRE"
   grd_LisCab.Col = 37:  grd_LisCab.Text = "DICIEMBRE"
   
   'Segunda linea
   grd_LisCab.Rows = grd_LisCab.Rows + 1
   grd_LisCab.Row = grd_LisCab.Rows - 1
   grd_LisCab.Col = 1:   grd_LisCab.Text = "CLASIFICACIONES"
   grd_LisCab.Col = 2:   grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 3:   grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 4:   grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 5:   grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 6:   grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 7:   grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 8:   grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 9:   grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 10:  grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 11:  grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 12:  grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 13:  grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 14:  grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 15:  grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 16:  grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 17:  grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 18:  grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 19:  grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 20:  grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 21:  grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 22:  grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 23:  grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 24:  grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 25:  grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 26:  grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 27:  grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 28:  grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 29:  grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 30:  grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 31:  grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 32:  grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 33:  grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 34:  grd_LisCab.Text = "PROVISION S/."
   grd_LisCab.Col = 35:  grd_LisCab.Text = "NUMERO"
   grd_LisCab.Col = 36:  grd_LisCab.Text = "MONTO S/."
   grd_LisCab.Col = 37:  grd_LisCab.Text = "PROVISION S/."
   
   'Tercera linea
   grd_LisCab.Rows = grd_LisCab.Rows + 1
   grd_LisCab.Row = grd_LisCab.Rows - 1
   grd_LisCab.Col = 0:   grd_LisCab.Text = "0"
   grd_LisCab.Col = 1:   grd_LisCab.Text = "NORMAL"
   
   'Cuarta linea
   grd_LisCab.Rows = grd_LisCab.Rows + 1
   grd_LisCab.Row = grd_LisCab.Rows - 1
   grd_LisCab.Col = 0:   grd_LisCab.Text = "1"
   grd_LisCab.Col = 1:   grd_LisCab.Text = "CPP"
   
   'Quinta linea
   grd_LisCab.Rows = grd_LisCab.Rows + 1
   grd_LisCab.Row = grd_LisCab.Rows - 1
   grd_LisCab.Col = 0:   grd_LisCab.Text = "2"
   grd_LisCab.Col = 1:   grd_LisCab.Text = "DEFICIENTE"
   
   'Sexta linea
   grd_LisCab.Rows = grd_LisCab.Rows + 1
   grd_LisCab.Row = grd_LisCab.Rows - 1
   grd_LisCab.Col = 0:   grd_LisCab.Text = "3"
   grd_LisCab.Col = 1:   grd_LisCab.Text = "DUDOSO"
   
   'Setima linea
   grd_LisCab.Rows = grd_LisCab.Rows + 1
   grd_LisCab.Row = grd_LisCab.Rows - 1
   grd_LisCab.Col = 0:   grd_LisCab.Text = "4"
   grd_LisCab.Col = 1:   grd_LisCab.Text = "PERDIDA"
   
   'Totales
   grd_LisCab.Rows = grd_LisCab.Rows + 2
   grd_LisCab.Row = grd_LisCab.Rows - 1
   grd_LisCab.Col = 0:   grd_LisCab.Text = ""
   grd_LisCab.Col = 1:   grd_LisCab.Text = "TOTALES"
   
   'Alineados
   grd_LisCab.Rows = grd_LisCab.Rows + 1
   grd_LisCab.Row = grd_LisCab.Rows - 1
   grd_LisCab.Col = 0:   grd_LisCab.Text = ""
   grd_LisCab.Col = 1:   grd_LisCab.Text = "ALINEADOS"
   
   With grd_LisCab
      .MergeCells = flexMergeFree
      .MergeCol(1) = True
      .MergeRow(0) = True
      .FixedCols = 2
      .FixedRows = 2
   End With
   
   For r_int_Contad = 1 To r_int_PerMes
      'Prepara SP que trae consolidado mensual
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT A.HIPCIE_CLAPRV AS CLASIFICACION, "
      g_str_Parame = g_str_Parame & "       COUNT(*) AS NUMERO, "
      g_str_Parame = g_str_Parame & "       SUM(DECODE(A.HIPCIE_TIPMON, 1, (A.HIPCIE_SALCAP + A.HIPCIE_SALCON), A.HIPCIE_TIPCAM * (A.HIPCIE_SALCAP + A.HIPCIE_SALCON))) AS MONTO_TOTAL, "
      g_str_Parame = g_str_Parame & "       SUM(DECODE(A.HIPCIE_TIPMON, 1, (A.HIPCIE_PRVGEN + A.HIPCIE_PRVESP + A.HIPCIE_PRVCIC + A.HIPCIE_PRVGEN_RC + A.HIPCIE_PRVCIC_RC + A.HIPCIE_PRVVOL), A.HIPCIE_TIPCAM*(A.HIPCIE_PRVGEN + A.HIPCIE_PRVESP + A.HIPCIE_PRVCIC + A.HIPCIE_PRVGEN_RC + A.HIPCIE_PRVCIC_RC + A.HIPCIE_PRVVOL))) AS PROVISIONES, "
      g_str_Parame = g_str_Parame & "       (SELECT COUNT(*) FROM CRE_HIPCIE B WHERE B.HIPCIE_PERMES = " & CStr(r_int_Contad) & " AND B.HIPCIE_PERANO = " & CStr(r_int_PerAno) & " AND B.HIPCIE_CLACLI <> B.HIPCIE_CLAALI AND B.HIPCIE_CLAALI > 2) AS TOT_NUM_ALI, "
      g_str_Parame = g_str_Parame & "       (SELECT SUM(DECODE(C.HIPCIE_TIPMON, 1, (C.HIPCIE_SALCAP + C.HIPCIE_SALCON), C.HIPCIE_TIPCAM * (C.HIPCIE_SALCAP + C.HIPCIE_SALCON))) FROM CRE_HIPCIE C WHERE C.HIPCIE_PERMES = " & CStr(r_int_Contad) & " AND C.HIPCIE_PERANO = " & CStr(r_int_PerAno) & " AND C.HIPCIE_CLACLI <> C.HIPCIE_CLAALI AND C.HIPCIE_CLAALI > 2) AS TOT_SAL_ALI, "
      g_str_Parame = g_str_Parame & "       (SELECT SUM(DECODE(D.HIPCIE_TIPMON, 1, (D.HIPCIE_PRVGEN + D.HIPCIE_PRVESP + D.HIPCIE_PRVCIC + D.HIPCIE_PRVGEN_RC + D.HIPCIE_PRVCIC_RC + D.HIPCIE_PRVVOL), D.HIPCIE_TIPCAM*(D.HIPCIE_PRVGEN + D.HIPCIE_PRVESP + D.HIPCIE_PRVCIC + D.HIPCIE_PRVGEN_RC + D.HIPCIE_PRVCIC_RC + D.HIPCIE_PRVVOL))) FROM CRE_HIPCIE D WHERE D.HIPCIE_PERMES = " & CStr(r_int_Contad) & " AND D.HIPCIE_PERANO = " & CStr(r_int_PerAno) & " AND D.HIPCIE_CLACLI <> D.HIPCIE_CLAALI AND D.HIPCIE_CLAALI > 2) AS TOT_PRV_ALI "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE A "
      g_str_Parame = g_str_Parame & " WHERE A.HIPCIE_PERANO = " & CStr(r_int_PerAno) & " "
      g_str_Parame = g_str_Parame & "   AND A.HIPCIE_PERMES = " & CStr(r_int_Contad) & " "
      If Not (r_int_Produc = 0) Then
         If (r_int_Produc = 1) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('001') "
         End If
         If (r_int_Produc = 2) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('002','006','011') "
         End If
         If (r_int_Produc = 3) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('003') "
         End If
         If (r_int_Produc = 4) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025') " ','019','021','022','023'
         End If
         If (r_int_Produc = 5) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('019') "
         End If
         If (r_int_Produc = 6) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('021','022','023') "
         End If
         If (r_int_Produc = 7) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('024') "
         End If
      End If
      g_str_Parame = g_str_Parame & "GROUP BY A.HIPCIE_CLAPRV "
      
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "Error al ejecutar la consulta de Consolidado de Clasificaciones.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      r_int_NumCol = (r_int_Contad * 3) - 1
      r_int_NumFil = 2
      r_int_TotNum = 0
      r_dbl_TotMto = 0
      r_dbl_TotPrv = 0
      r_int_NumAli = 0
      r_dbl_TotAli = 0
      r_dbl_PrvAli = 0
      
      'Carga grid
      For r_int_ConCla = 0 To 4
         r_int_Numero = 0
         r_dbl_MtoTot = 0
         r_dbl_MtoPrv = 0
         Call fs_Obtiene_DatosClasificacionCab(r_int_ConCla, r_int_Numero, r_dbl_MtoTot, r_dbl_MtoPrv, r_int_NumAli, r_dbl_TotAli, r_dbl_PrvAli)
         
         grd_LisCab.Col = r_int_NumCol
         grd_LisCab.Row = r_int_NumFil
         grd_LisCab.Text = Format(r_int_Numero, "##,##0")
         r_int_TotNum = r_int_TotNum + r_int_Numero
         
         grd_LisCab.Col = r_int_NumCol + 1
         grd_LisCab.Row = r_int_NumFil
         grd_LisCab.Text = Format(r_dbl_MtoTot, "###,###,##.00")
         r_dbl_TotMto = r_dbl_TotMto + r_dbl_MtoTot
         
         grd_LisCab.Col = r_int_NumCol + 2
         grd_LisCab.Row = r_int_NumFil
         grd_LisCab.Text = Format(r_dbl_MtoPrv, "###,###,##.00")
         r_dbl_TotPrv = r_dbl_TotPrv + r_dbl_MtoPrv
         
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_ConCla
      
      'Carga Totales
      grd_LisCab.Col = r_int_NumCol
      grd_LisCab.Row = r_int_NumFil + 1
      grd_LisCab.Text = Format(r_int_TotNum, "##,##0")
      
      grd_LisCab.Col = r_int_NumCol + 1
      grd_LisCab.Row = r_int_NumFil + 1
      grd_LisCab.Text = Format(r_dbl_TotMto, "###,###,##.00")
      
      grd_LisCab.Col = r_int_NumCol + 2
      grd_LisCab.Row = r_int_NumFil + 1
      grd_LisCab.Text = Format(r_dbl_TotPrv, "###,###,##.00")
      
      'Carga Alineados
      grd_LisCab.Col = r_int_NumCol
      grd_LisCab.Row = r_int_NumFil + 2
      grd_LisCab.Text = Format(r_int_NumAli, "##,##0")
      
      grd_LisCab.Col = r_int_NumCol + 1
      grd_LisCab.Row = r_int_NumFil + 2
      grd_LisCab.Text = Format(r_dbl_TotAli, "###,###,##.00")
      
      grd_LisCab.Col = r_int_NumCol + 2
      grd_LisCab.Row = r_int_NumFil + 2
      grd_LisCab.Text = Format(r_dbl_PrvAli, "###,###,##.00")
            
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next
   
   grd_LisCab.Redraw = True
   grd_LisCab.Enabled = True
   Call gs_UbicaGrid(grd_LisCab, 2)
End Sub

Private Sub fs_Obtiene_Detalle()
Dim r_int_Contad     As Integer
Dim r_int_ConCla     As Integer
Dim r_int_NumCol     As Integer
Dim r_int_NumFil     As Integer
Dim r_bol_DetCla     As Boolean

Dim r_int_NumReg     As Integer
Dim r_dbl_MtoTot     As Double
Dim r_dbl_MtoPrv     As Double
Dim r_int_CarNum     As Integer
Dim r_dbl_CarTot     As Double
Dim r_dbl_CarPrv     As Double
Dim r_int_AliNum     As Integer
Dim r_dbl_AliTot     As Double
Dim r_dbl_AliPrv     As Double
Dim r_int_TotNum     As Integer
Dim r_dbl_TotMto     As Double
Dim r_dbl_TotPrv     As Double

   grd_LisDet.Redraw = False
   Call gs_LimpiaGrid(grd_LisDet)
   
   'Primera Linea
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Row = 0:   grd_LisDet.Text = ""
   grd_LisDet.Col = 1:   grd_LisDet.Text = "CLASIFICACIONES"
   grd_LisDet.Col = 2:   grd_LisDet.Text = "ENERO"
   grd_LisDet.Col = 3:   grd_LisDet.Text = "ENERO"
   grd_LisDet.Col = 4:   grd_LisDet.Text = "ENERO"
   grd_LisDet.Col = 5:   grd_LisDet.Text = "FEBRERO"
   grd_LisDet.Col = 6:   grd_LisDet.Text = "FEBRERO"
   grd_LisDet.Col = 7:   grd_LisDet.Text = "FEBRERO"
   grd_LisDet.Col = 8:   grd_LisDet.Text = "MARZO"
   grd_LisDet.Col = 9:   grd_LisDet.Text = "MARZO"
   grd_LisDet.Col = 10:  grd_LisDet.Text = "MARZO"
   grd_LisDet.Col = 11:  grd_LisDet.Text = "ABRIL"
   grd_LisDet.Col = 12:  grd_LisDet.Text = "ABRIL"
   grd_LisDet.Col = 13:  grd_LisDet.Text = "ABRIL"
   grd_LisDet.Col = 14:  grd_LisDet.Text = "MAYO"
   grd_LisDet.Col = 15:  grd_LisDet.Text = "MAYO"
   grd_LisDet.Col = 16:  grd_LisDet.Text = "MAYO"
   grd_LisDet.Col = 17:  grd_LisDet.Text = "JUNIO"
   grd_LisDet.Col = 18:  grd_LisDet.Text = "JUNIO"
   grd_LisDet.Col = 19:  grd_LisDet.Text = "JUNIO"
   grd_LisDet.Col = 20:  grd_LisDet.Text = "JULIO"
   grd_LisDet.Col = 21:  grd_LisDet.Text = "JULIO"
   grd_LisDet.Col = 22:  grd_LisDet.Text = "JULIO"
   grd_LisDet.Col = 23:  grd_LisDet.Text = "AGOSTO"
   grd_LisDet.Col = 24:  grd_LisDet.Text = "AGOSTO"
   grd_LisDet.Col = 25:  grd_LisDet.Text = "AGOSTO"
   grd_LisDet.Col = 26:  grd_LisDet.Text = "SETIEMBRE"
   grd_LisDet.Col = 27:  grd_LisDet.Text = "SETIEMBRE"
   grd_LisDet.Col = 28:  grd_LisDet.Text = "SETIEMBRE"
   grd_LisDet.Col = 29:  grd_LisDet.Text = "OCTUBRE"
   grd_LisDet.Col = 30:  grd_LisDet.Text = "OCTUBRE"
   grd_LisDet.Col = 31:  grd_LisDet.Text = "OCTUBRE"
   grd_LisDet.Col = 32:  grd_LisDet.Text = "NOVIEMBRE"
   grd_LisDet.Col = 33:  grd_LisDet.Text = "NOVIEMBRE"
   grd_LisDet.Col = 34:  grd_LisDet.Text = "NOVIEMBRE"
   grd_LisDet.Col = 35:  grd_LisDet.Text = "DICIEMBRE"
   grd_LisDet.Col = 36:  grd_LisDet.Text = "DICIEMBRE"
   grd_LisDet.Col = 37:  grd_LisDet.Text = "DICIEMBRE"
   
   'Segunda linea
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 1:   grd_LisDet.Text = "CLASIFICACIONES"
   grd_LisDet.Col = 2:   grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 3:   grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 4:   grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 5:   grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 6:   grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 7:   grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 8:   grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 9:   grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 10:  grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 11:  grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 12:  grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 13:  grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 14:  grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 15:  grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 16:  grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 17:  grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 18:  grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 19:  grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 20:  grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 21:  grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 22:  grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 23:  grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 24:  grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 25:  grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 26:  grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 27:  grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 28:  grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 29:  grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 30:  grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 31:  grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 32:  grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 33:  grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 34:  grd_LisDet.Text = "PROVISION S/."
   grd_LisDet.Col = 35:  grd_LisDet.Text = "NUMERO"
   grd_LisDet.Col = 36:  grd_LisDet.Text = "MONTO S/."
   grd_LisDet.Col = 37:  grd_LisDet.Text = "PROVISION S/."
   
   'Tercera linea
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "0"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "NORMAL"
   
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "0"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "       CARTERA"
   
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "0"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "       ALIENADOS"
   
   'Cuarta linea
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "1"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "CPP"
   
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "1"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "       CARTERA"
   
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "1"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "       ALIENADOS"
   
   'Quinta linea
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "2"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "DEFICIENTE"
   
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "2"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "       CARTERA"
   
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "2"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "       ALIENADOS"
   
   'Sexta linea
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "3"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "DUDOSO"
   
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "3"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "       CARTERA"
   
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "3"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "       ALIENADOS"
   
   'Setima linea
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "4"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "PERDIDA"
   
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "4"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "       CARTERA"
   
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = "4"
   grd_LisDet.Col = 1:   grd_LisDet.Text = "       ALIENADOS"
   
   'Total Cartera
   grd_LisDet.Rows = grd_LisDet.Rows + 2
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = ""
   grd_LisDet.Col = 1:   grd_LisDet.Text = "TOTAL CARTERA"
   
   'Total Alineados
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = ""
   grd_LisDet.Col = 1:   grd_LisDet.Text = "TOTAL ALINEADOS"
   
   'Total General
   grd_LisDet.Rows = grd_LisDet.Rows + 1
   grd_LisDet.Row = grd_LisDet.Rows - 1
   grd_LisDet.Col = 0:   grd_LisDet.Text = ""
   grd_LisDet.Col = 1:   grd_LisDet.Text = "TOTAL GENERAL"
   
   With grd_LisDet
      .MergeCells = flexMergeFree
      .MergeCol(1) = True
      .MergeRow(0) = True
      .FixedCols = 2
      .FixedRows = 2
   End With
   
   For r_int_Contad = 1 To r_int_PerMes
      'Prepara SP que trae consolidado mensual
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT A.HIPCIE_CLAPRV AS CLASIFICACION, "
      g_str_Parame = g_str_Parame & "       COUNT(*) AS NUMERO, "
      g_str_Parame = g_str_Parame & "       ROUND(SUM(DECODE(A.HIPCIE_TIPMON, 1, (A.HIPCIE_SALCAP + A.HIPCIE_SALCON), A.HIPCIE_TIPCAM * (A.HIPCIE_SALCAP + A.HIPCIE_SALCON))),2) AS MONTO_TOTAL, "
      g_str_Parame = g_str_Parame & "       ROUND(SUM(DECODE(A.HIPCIE_TIPMON, 1, (A.HIPCIE_PRVGEN + A.HIPCIE_PRVESP + A.HIPCIE_PRVCIC + A.HIPCIE_PRVGEN_RC + A.HIPCIE_PRVCIC_RC + A.HIPCIE_PRVVOL), A.HIPCIE_TIPCAM*(A.HIPCIE_PRVGEN + A.HIPCIE_PRVESP + A.HIPCIE_PRVCIC + A.HIPCIE_PRVGEN_RC + A.HIPCIE_PRVCIC_RC + A.HIPCIE_PRVVOL))),2) AS PROVISIONES, "
      g_str_Parame = g_str_Parame & "       (SELECT COUNT(*) FROM CRE_HIPCIE B WHERE B.HIPCIE_PERMES = " & CStr(r_int_Contad) & " AND B.HIPCIE_PERANO = " & CStr(r_int_PerAno) & " AND B.HIPCIE_CLACLI <> B.HIPCIE_CLAALI AND B.HIPCIE_CLAALI > 2) AS TOT_NUM_ALI, "
      g_str_Parame = g_str_Parame & "       (SELECT SUM(DECODE(C.HIPCIE_TIPMON, 1, (C.HIPCIE_SALCAP + C.HIPCIE_SALCON), C.HIPCIE_TIPCAM * (C.HIPCIE_SALCAP + C.HIPCIE_SALCON))) FROM CRE_HIPCIE C WHERE C.HIPCIE_PERMES = " & CStr(r_int_Contad) & " AND C.HIPCIE_PERANO = " & CStr(r_int_PerAno) & " AND C.HIPCIE_CLACLI <> C.HIPCIE_CLAALI AND C.HIPCIE_CLAALI > 2) AS TOT_SAL_ALI, "
      g_str_Parame = g_str_Parame & "       (SELECT SUM(DECODE(D.HIPCIE_TIPMON, 1, (D.HIPCIE_PRVGEN + D.HIPCIE_PRVESP + D.HIPCIE_PRVCIC + D.HIPCIE_PRVGEN_RC + D.HIPCIE_PRVCIC_RC + D.HIPCIE_PRVVOL ), D.HIPCIE_TIPCAM*(D.HIPCIE_PRVGEN + D.HIPCIE_PRVESP + D.HIPCIE_PRVCIC + D.HIPCIE_PRVGEN_RC + D.HIPCIE_PRVCIC_RC + D.HIPCIE_PRVVOL))) FROM CRE_HIPCIE D WHERE D.HIPCIE_PERMES = " & CStr(r_int_Contad) & " AND D.HIPCIE_PERANO = " & CStr(r_int_PerAno) & " AND D.HIPCIE_CLACLI <> D.HIPCIE_CLAALI AND D.HIPCIE_CLAALI > 2) AS TOT_PRV_ALI "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE A "
      g_str_Parame = g_str_Parame & " WHERE A.HIPCIE_PERANO = " & CStr(r_int_PerAno) & " "
      g_str_Parame = g_str_Parame & "   AND A.HIPCIE_PERMES = " & CStr(r_int_Contad) & " "
      If Not (r_int_Produc = 0) Then
         If (r_int_Produc = 1) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('001') "
         End If
         If (r_int_Produc = 2) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('002','006','011') " ','012'
         End If
         If (r_int_Produc = 3) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('003') "
         End If
         If (r_int_Produc = 4) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025') " ','019'
         End If
         If (r_int_Produc = 5) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('019') "
         End If
         If (r_int_Produc = 6) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('021','022','023') "
         End If
         If (r_int_Produc = 7) Then
            g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CODPRD IN ('024') "
         End If
      End If
      g_str_Parame = g_str_Parame & "GROUP BY A.HIPCIE_CLAPRV "
      
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "Error al ejecutar la consulta de Consolidado de Clasificaciones.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      r_int_NumCol = (r_int_Contad * 3) - 1
      r_int_NumFil = 2
      r_int_CarNum = 0
      r_dbl_CarTot = 0
      r_dbl_CarPrv = 0
      r_int_AliNum = 0
      r_dbl_AliTot = 0
      r_dbl_AliPrv = 0
      r_int_TotNum = 0
      r_dbl_TotMto = 0
      r_dbl_TotPrv = 0
      
      'Detalle por Cartera y Alienados
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT A.HIPCIE_CLAPRV AS CLASIFICACION, "
      g_str_Parame = g_str_Parame & "       COUNT(*) AS NUMERO, "
      g_str_Parame = g_str_Parame & "       ROUND(SUM(DECODE(A.HIPCIE_TIPMON, 1, (A.HIPCIE_SALCAP + A.HIPCIE_SALCON), A.HIPCIE_TIPCAM * (A.HIPCIE_SALCAP + A.HIPCIE_SALCON))),2) AS MONTO_TOTAL, "
      g_str_Parame = g_str_Parame & "       ROUND(SUM(DECODE(A.HIPCIE_TIPMON, 1, (A.HIPCIE_PRVGEN + A.HIPCIE_PRVESP + A.HIPCIE_PRVCIC + A.HIPCIE_PRVGEN_RC + A.HIPCIE_PRVCIC_RC + A.HIPCIE_PRVVOL), A.HIPCIE_TIPCAM*(A.HIPCIE_PRVGEN + A.HIPCIE_PRVESP + A.HIPCIE_PRVCIC + A.HIPCIE_PRVGEN_RC + A.HIPCIE_PRVCIC_RC + A.HIPCIE_PRVVOL))),2) AS PROVISIONES "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE A "
      g_str_Parame = g_str_Parame & " WHERE A.HIPCIE_PERANO = " & CStr(r_int_PerAno) & " "
      g_str_Parame = g_str_Parame & "   AND A.HIPCIE_PERMES = " & CStr(r_int_Contad) & " "
      g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CLACLI <> HIPCIE_CLAALI "
      g_str_Parame = g_str_Parame & "   AND A.HIPCIE_CLAALI  > 2 "
      g_str_Parame = g_str_Parame & " GROUP BY HIPCIE_CLAPRV "
      
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         MsgBox "Error al ejecutar la consulta de Detalle de Clasificaciones.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      'Carga grid
      For r_int_ConCla = 0 To 4
         r_int_NumReg = 0
         r_dbl_MtoTot = 0
         r_dbl_MtoPrv = 0
         r_bol_DetCla = False
         Call fs_Obtiene_DatosClasificacion(r_int_ConCla, r_int_NumReg, r_dbl_MtoTot, r_dbl_MtoPrv)
         
         'Acumulado por Clasificacion
         grd_LisDet.Col = r_int_NumCol
         grd_LisDet.Row = r_int_NumFil
         grd_LisDet.Text = Format(r_int_NumReg, "##,##0")
         r_int_TotNum = r_int_TotNum + r_int_NumReg
         
         grd_LisDet.Col = r_int_NumCol + 1
         grd_LisDet.Row = r_int_NumFil
         grd_LisDet.Text = Format(r_dbl_MtoTot, "###,###,##.00")
         r_dbl_TotMto = r_dbl_TotMto + r_dbl_MtoTot
         
         grd_LisDet.Col = r_int_NumCol + 2
         grd_LisDet.Row = r_int_NumFil
         grd_LisDet.Text = Format(r_dbl_MtoPrv, "###,###,##.00")
         r_dbl_TotPrv = r_dbl_TotPrv + r_dbl_MtoPrv
         
         'Busca Detalle
         If Not (g_rst_Genera.EOF And g_rst_Genera.BOF) Then
            g_rst_Genera.MoveFirst
            Do While Not g_rst_Genera.EOF
               'Compara Clasificaciones
               If r_int_ConCla = g_rst_Genera!CLASIFICACION Then
                  r_bol_DetCla = True
                  r_int_NumFil = r_int_NumFil + 1
                  grd_LisDet.Col = r_int_NumCol
                  grd_LisDet.Row = r_int_NumFil
                  r_int_CarNum = r_int_CarNum + (r_int_NumReg - g_rst_Genera!numero)
                  grd_LisDet.Text = Format(r_int_NumReg - g_rst_Genera!numero, "##,##0")
                  
                  grd_LisDet.Col = r_int_NumCol + 1
                  grd_LisDet.Row = r_int_NumFil
                  r_dbl_CarTot = r_dbl_CarTot + (r_dbl_MtoTot - g_rst_Genera!MONTO_TOTAL)
                  grd_LisDet.Text = Format(r_dbl_MtoTot - g_rst_Genera!MONTO_TOTAL, "###,###,##.00")
                  
                  grd_LisDet.Col = r_int_NumCol + 2
                  grd_LisDet.Row = r_int_NumFil
                  If Not IsNull(g_rst_Genera!PROVISIONES) Then
                     r_dbl_CarPrv = r_dbl_CarPrv + (r_dbl_MtoPrv - g_rst_Genera!PROVISIONES)
                  End If
                  grd_LisDet.Text = Format(r_dbl_MtoPrv - g_rst_Genera!PROVISIONES, "###,###,##.00")
                  
                  r_int_NumFil = r_int_NumFil + 1
                  grd_LisDet.Col = r_int_NumCol
                  grd_LisDet.Row = r_int_NumFil
                  r_int_AliNum = r_int_AliNum + g_rst_Genera!numero
                  grd_LisDet.Text = Format(g_rst_Genera!numero, "##,##0")
                  
                  grd_LisDet.Col = r_int_NumCol + 1
                  grd_LisDet.Row = r_int_NumFil
                  r_dbl_AliTot = r_dbl_AliTot + g_rst_Genera!MONTO_TOTAL
                  grd_LisDet.Text = Format(g_rst_Genera!MONTO_TOTAL, "###,###,##.00")
                  
                  grd_LisDet.Col = r_int_NumCol + 2
                  grd_LisDet.Row = r_int_NumFil
                  If Not IsNull(g_rst_Genera!PROVISIONES) Then
                     r_dbl_AliPrv = r_dbl_AliPrv + g_rst_Genera!PROVISIONES
                  End If
                  grd_LisDet.Text = Format(g_rst_Genera!PROVISIONES, "###,###,##.00")
               End If
               
               g_rst_Genera.MoveNext
            Loop
         End If
         
         If Not r_bol_DetCla Then
            r_int_NumFil = r_int_NumFil + 1
            grd_LisDet.Col = r_int_NumCol
            grd_LisDet.Row = r_int_NumFil
            grd_LisDet.Text = Format(r_int_NumReg, "##,##0")
            r_int_CarNum = r_int_CarNum + r_int_NumReg
            
            grd_LisDet.Col = r_int_NumCol + 1
            grd_LisDet.Row = r_int_NumFil
            grd_LisDet.Text = Format(r_dbl_MtoTot, "###,###,##.00")
            r_dbl_CarTot = r_dbl_CarTot + r_dbl_MtoTot
            
            grd_LisDet.Col = r_int_NumCol + 2
            grd_LisDet.Row = r_int_NumFil
            grd_LisDet.Text = Format(r_dbl_MtoPrv, "###,###,##.00")
            r_dbl_CarPrv = r_dbl_CarPrv + r_dbl_MtoPrv
            
            r_int_NumFil = r_int_NumFil + 1
            grd_LisDet.Col = r_int_NumCol
            grd_LisDet.Row = r_int_NumFil
            grd_LisDet.Text = Format(0, "##,##0")
            
            grd_LisDet.Col = r_int_NumCol + 1
            grd_LisDet.Row = r_int_NumFil
            grd_LisDet.Text = Format(0, "###,###,##.00")
            
            grd_LisDet.Col = r_int_NumCol + 2
            grd_LisDet.Row = r_int_NumFil
            grd_LisDet.Text = Format(0, "###,###,##.00")
         End If
         
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_ConCla
      
      'Carga Cartera
      grd_LisDet.Col = r_int_NumCol
      grd_LisDet.Row = r_int_NumFil + 1
      grd_LisDet.Text = Format(r_int_CarNum, "##,##0")
      
      grd_LisDet.Col = r_int_NumCol + 1
      grd_LisDet.Row = r_int_NumFil + 1
      grd_LisDet.Text = Format(r_dbl_CarTot, "###,###,##.00")
      
      grd_LisDet.Col = r_int_NumCol + 2
      grd_LisDet.Row = r_int_NumFil + 1
      grd_LisDet.Text = Format(r_dbl_CarPrv, "###,###,##.00")
      
      'Carga Alineados
      grd_LisDet.Col = r_int_NumCol
      grd_LisDet.Row = r_int_NumFil + 2
      grd_LisDet.Text = Format(r_int_AliNum, "##,##0")
      
      grd_LisDet.Col = r_int_NumCol + 1
      grd_LisDet.Row = r_int_NumFil + 2
      grd_LisDet.Text = Format(r_dbl_AliTot, "###,###,##.00")
      
      grd_LisDet.Col = r_int_NumCol + 2
      grd_LisDet.Row = r_int_NumFil + 2
      grd_LisDet.Text = Format(r_dbl_AliPrv, "###,###,##.00")
      
      'Carga Totales
      grd_LisDet.Col = r_int_NumCol
      grd_LisDet.Row = r_int_NumFil + 3
      grd_LisDet.Text = Format(r_int_CarNum + r_int_AliNum, "##,##0")
      
      grd_LisDet.Col = r_int_NumCol + 1
      grd_LisDet.Row = r_int_NumFil + 3
      grd_LisDet.Text = Format(r_dbl_CarTot + r_dbl_AliTot, "###,###,##.00")
      
      grd_LisDet.Col = r_int_NumCol + 2
      grd_LisDet.Row = r_int_NumFil + 3
      grd_LisDet.Text = Format(r_dbl_CarPrv + r_dbl_AliPrv, "###,###,##.00")
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next
   
   grd_LisDet.Redraw = True
   grd_LisDet.Enabled = True
   Call gs_UbicaGrid(grd_LisDet, 2)
End Sub

Private Sub fs_Obtiene_DatosClasificacionCab(ByVal p_CodCla As Integer, ByRef p_Numero As Integer, ByRef p_MtoTot As Double, ByRef p_MtoPrv As Double, ByRef p_NumAli As Integer, ByRef p_TotAli As Double, ByRef p_PrvAli As Double)
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If p_CodCla = g_rst_Princi!CLASIFICACION Then
            If Not IsNull(g_rst_Princi!numero) Then
               p_Numero = g_rst_Princi!numero
            End If
            If Not IsNull(g_rst_Princi!MONTO_TOTAL) Then
               p_MtoTot = g_rst_Princi!MONTO_TOTAL
            End If
            If Not IsNull(g_rst_Princi!PROVISIONES) Then
               p_MtoPrv = g_rst_Princi!PROVISIONES
            End If
            If Not IsNull(g_rst_Princi!TOT_NUM_ALI) Then
               p_NumAli = g_rst_Princi!TOT_NUM_ALI
            End If
            If Not IsNull(g_rst_Princi!TOT_SAL_ALI) Then
               p_TotAli = g_rst_Princi!TOT_SAL_ALI
            End If
            If Not IsNull(g_rst_Princi!TOT_PRV_ALI) Then
               p_PrvAli = g_rst_Princi!TOT_PRV_ALI
            End If
         End If
         g_rst_Princi.MoveNext
      Loop
   End If
End Sub

Private Sub fs_Obtiene_DatosClasificacion(ByVal p_CodCla As Integer, ByRef p_Numero As Integer, ByRef p_MtoTot As Double, ByRef p_MtoPrv As Double)
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If p_CodCla = g_rst_Princi!CLASIFICACION Then
            If Not IsNull(g_rst_Princi!numero) Then
               p_Numero = g_rst_Princi!numero
            End If
            If Not IsNull(g_rst_Princi!MONTO_TOTAL) Then
               p_MtoTot = g_rst_Princi!MONTO_TOTAL
            End If
            If Not IsNull(g_rst_Princi!PROVISIONES) Then
               p_MtoPrv = g_rst_Princi!PROVISIONES
            End If
         End If
         g_rst_Princi.MoveNext
      Loop
   End If
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_Produc.ListIndex > -1 Then
         Call gs_SetFocus(cmb_PerMes)
      End If
   End If
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If CInt(ipp_PerAno.Text) >= 2007 Then
         Call gs_SetFocus(cmd_Proces)
      End If
   End If
End Sub

Private Sub grd_LisCab_Click()
   If grd_LisCab.Rows = 0 Then
      Exit Sub
   End If

   moddat_g_int_TipCli = 1
   moddat_g_int_EdaMes = 0
   moddat_g_int_EdaAno = 0
   moddat_g_str_TipPar = ""
   moddat_g_str_NomPrd = ""
   
   moddat_g_int_EdaMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   moddat_g_int_EdaAno = CInt(ipp_PerAno.Text)
   moddat_g_str_TipPar = Trim(grd_LisCab.TextMatrix(grd_LisCab.Row, 0))
   moddat_g_str_NomPrd = " '" & Trim(grd_LisCab.TextMatrix(grd_LisCab.Row, 1)) & "' A " & Trim(cmb_PerMes.Text) & " DEL " & CStr(ipp_PerAno.Text)
   
   If (moddat_g_int_EdaMes > 0) And (moddat_g_int_EdaAno > 0) And (Len(Trim(moddat_g_str_TipPar)) > 0) And Len(Trim(moddat_g_str_NomPrd)) > 0 Then
      frm_RptCtb_25.Show 1
   End If
End Sub

Private Sub grd_LisDet_DblClick()
   If grd_LisDet.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_TipCli = 1
   moddat_g_int_EdaMes = 0
   moddat_g_int_EdaAno = 0
   moddat_g_str_TipPar = ""
   moddat_g_str_NomPrd = ""
   
   moddat_g_int_EdaMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   moddat_g_int_EdaAno = CInt(ipp_PerAno.Text)
   moddat_g_str_TipPar = Trim(grd_LisDet.TextMatrix(grd_LisDet.Row, 0))
   Select Case moddat_g_str_TipPar
      Case 0: moddat_g_str_NomPrd = " 'NORMAL' A " & Trim(cmb_PerMes.Text) & " DEL " & CStr(ipp_PerAno.Text)
      Case 1: moddat_g_str_NomPrd = " 'CPP' A " & Trim(cmb_PerMes.Text) & " DEL " & CStr(ipp_PerAno.Text)
      Case 2: moddat_g_str_NomPrd = " 'DEFICIENTE' A " & Trim(cmb_PerMes.Text) & " DEL " & CStr(ipp_PerAno.Text)
      Case 3: moddat_g_str_NomPrd = " 'DUDOSO' A " & Trim(cmb_PerMes.Text) & " DEL " & CStr(ipp_PerAno.Text)
      Case 4: moddat_g_str_NomPrd = " 'PERDIDA' A " & Trim(cmb_PerMes.Text) & " DEL " & CStr(ipp_PerAno.Text)
   End Select
   
   If (moddat_g_int_EdaMes > 0) And (moddat_g_int_EdaAno > 0) And (Len(Trim(moddat_g_str_TipPar)) > 0) And Len(Trim(moddat_g_str_NomPrd)) > 0 Then
      frm_RptCtb_25.Show 1
   End If
End Sub

Private Sub fs_GenExcRes()
Dim r_obj_Excel      As Excel.Application
Dim r_int_Contad     As Integer
Dim r_int_NroFil     As Integer
Dim r_int_NoFlLi     As Integer
Dim r_int_TotReg     As Integer

   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      'Titulo
      If CInt(tab_Clasif.Tab) = 2 Then
         .Cells(2, 2) = "CLASIFICACION DE CARTERA (NORMALES, ATRASADOS Y ALINEADOS) DEL MES " & UCase(Trim(cmb_PerMes.Text)) & " DEL " & CStr(ipp_PerAno.Text)
         .Range(.Cells(2, 2), .Cells(2, 13)).Merge
         .Range("B2:M2").HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(2, 2), .Cells(2, 13)).Font.Name = "Calibri"
         .Range(.Cells(2, 2), .Cells(2, 13)).Font.Size = 12
         .Range(.Cells(2, 2), .Cells(2, 13)).Font.Bold = True
                  
         .Columns("C").NumberFormat = "###,##0"
         .Columns("D").NumberFormat = "###,##0"
         .Columns("E").NumberFormat = "###,##0"
         .Columns("F").NumberFormat = "###,##0"
         .Columns("G").NumberFormat = "###,##0"
         .Columns("H").NumberFormat = "###,##0"
         .Columns("I").NumberFormat = "###,##0"
         .Columns("J").NumberFormat = "###,##0"
         .Columns("K").NumberFormat = "###,##0"
         .Columns("L").NumberFormat = "###,##0"
         .Columns("M").NumberFormat = "###,##0"
         
         .Columns("B").ColumnWidth = 20
         .Columns("D").ColumnWidth = 15
         .Columns("F").ColumnWidth = 15
         .Columns("H").ColumnWidth = 15
         .Columns("J").ColumnWidth = 15
         .Columns("L").ColumnWidth = 15
         .Columns("M").ColumnWidth = 15
         
         .Range(.Cells(4, 2), .Cells(7, 2)).Merge
         .Range(.Cells(4, 2), .Cells(7, 2)) = "Producto"
         .Range(.Cells(4, 2), .Cells(7, 2)).WrapText = True
         .Range(.Cells(4, 2), .Cells(7, 2)).VerticalAlignment = xlCenter
         
         .Range(.Cells(4, 3), .Cells(4, 12)) = "Calificacion"
         .Range(.Cells(4, 3), .Cells(4, 12)).Merge
         .Range(.Cells(5, 3), .Cells(5, 4)) = "Normal"
         .Range(.Cells(5, 3), .Cells(5, 4)).Merge
         .Range(.Cells(5, 5), .Cells(5, 6)) = "CPP"
         .Range(.Cells(5, 5), .Cells(5, 6)).Merge
         .Range(.Cells(5, 7), .Cells(5, 8)) = "Deficiente"
         .Range(.Cells(5, 7), .Cells(5, 8)).Merge
         .Range(.Cells(5, 9), .Cells(5, 10)) = "Dudoso"
         .Range(.Cells(5, 9), .Cells(5, 10)).Merge
         .Range(.Cells(5, 11), .Cells(5, 12)) = "Perdida"
         .Range(.Cells(5, 11), .Cells(5, 12)).Merge
         
         .Range(.Cells(4, 13), .Cells(7, 13)).Merge
         .Range(.Cells(4, 13), .Cells(7, 13)) = "Total"
         .Range(.Cells(4, 13), .Cells(7, 13)).VerticalAlignment = xlCenter
         
         .Range(.Cells(6, 3), .Cells(6, 4)) = "01 - 30 dias"
         .Range(.Cells(6, 3), .Cells(6, 4)).Merge
         .Range(.Cells(6, 5), .Cells(6, 6)) = "31 - 60 dias"
         .Range(.Cells(6, 5), .Cells(6, 6)).Merge
         .Range(.Cells(6, 7), .Cells(6, 8)) = "61 - 120 dias"
         .Range(.Cells(6, 7), .Cells(6, 8)).Merge
         .Range(.Cells(6, 9), .Cells(6, 10)) = "121 - 365 dias"
         .Range(.Cells(6, 9), .Cells(6, 10)).Merge
         .Range(.Cells(6, 11), .Cells(6, 12)) = "mas de 365 dias"
         .Range(.Cells(6, 11), .Cells(6, 12)).Merge
         
'         r_obj_Excel.Visible = True
         
         For r_int_Contad = 0 To 8
            .Range(.Cells(7, r_int_Contad + 3), .Cells(7, r_int_Contad + 3)) = "N° Creditos"
            .Range(.Cells(7, r_int_Contad + 4), .Cells(7, r_int_Contad + 4)) = "Saldo"
            r_int_Contad = r_int_Contad + 1
         Next
                 
         
         .Range(.Cells(4, 2), .Cells(7, 13)).HorizontalAlignment = xlVAlignCenter
         .Range(.Cells(4, 2), .Cells(7, 13)).Font.Name = "Calibri"
         .Range(.Cells(4, 2), .Cells(7, 13)).Font.Size = 10
         .Range(.Cells(4, 2), .Cells(7, 13)).Font.Bold = True

         
         .Range(.Cells(4, 2), .Cells(7, 13)).Interior.Color = RGB(146, 208, 80)
         .Range(.Cells(4, 2), .Cells(7, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(5, 3), .Cells(5, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(6, 3), .Cells(6, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(7, 3), .Cells(7, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(4, 2), .Cells(8, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(4, 2), .Cells(7, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(4, 2), .Cells(7, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(4, 2), .Cells(7, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
         
         .Range(.Cells(8, 2), .Cells(9, 2)).Merge
         .Range(.Cells(8, 2), .Cells(9, 2)) = "CRC-PBP"
         .Range(.Cells(8, 2), .Cells(9, 2)).VerticalAlignment = xlCenter
         
         .Range(.Cells(10, 2), .Cells(11, 2)).Merge
         .Range(.Cells(10, 2), .Cells(11, 2)) = "Micasita"
         .Range(.Cells(10, 2), .Cells(11, 2)).VerticalAlignment = xlCenter
         
         .Range(.Cells(12, 2), .Cells(13, 2)).Merge
         .Range(.Cells(12, 2), .Cells(13, 2)) = "CME"
         .Range(.Cells(12, 2), .Cells(13, 2)).VerticalAlignment = xlCenter

         .Range(.Cells(14, 2), .Cells(15, 2)).Merge
         .Range(.Cells(14, 2), .Cells(15, 2)) = "N. MiVivienda"
         .Range(.Cells(14, 2), .Cells(15, 2)).VerticalAlignment = xlCenter

         .Range(.Cells(16, 2), .Cells(17, 2)).Merge
         .Range(.Cells(16, 2), .Cells(17, 2)) = "MiCasa Mas"
         .Range(.Cells(16, 2), .Cells(17, 2)).VerticalAlignment = xlCenter

         .Range(.Cells(18, 2), .Cells(19, 2)).Merge
         .Range(.Cells(18, 2), .Cells(19, 2)) = "MiVivienda Mas"
         .Range(.Cells(18, 2), .Cells(19, 2)).VerticalAlignment = xlCenter

         .Range(.Cells(20, 2), .Cells(21, 2)).Merge
         .Range(.Cells(20, 2), .Cells(21, 2)) = "BBP"
         .Range(.Cells(20, 2), .Cells(21, 2)).VerticalAlignment = xlCenter

         .Range(.Cells(22, 2), .Cells(23, 2)).Merge
         .Range(.Cells(22, 2), .Cells(23, 2)) = "Promedio Ponderado"
         .Range(.Cells(22, 2), .Cells(23, 2)).VerticalAlignment = xlCenter
                  
         .Range(.Cells(24, 2), .Cells(24, 2)) = "Total"
         .Range(.Cells(24, 2), .Cells(24, 2)).VerticalAlignment = xlCenter
         
         'CRC-PBP
         g_str_Parame = ""
         g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
         g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD = '001'"
         g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV "
'         g_str_Parame = g_str_Parame + " UNION SELECT DISTINCT HIPCIE_CLAPRV, 0, 0 FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         Do While Not g_rst_Princi.EOF
            If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
               .Range(.Cells(8, 3), .Cells(8, 3)) = g_rst_Princi!cont
               .Range(.Cells(8, 3), .Cells(8, 3)).VerticalAlignment = xlCenter
               .Range(.Cells(8, 4), .Cells(8, 4)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(8, 4), .Cells(8, 4)).VerticalAlignment = xlCenter
            End If
                        
            If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
               .Range(.Cells(8, 5), .Cells(8, 5)) = g_rst_Princi!cont
               .Range(.Cells(8, 5), .Cells(8, 5)).VerticalAlignment = xlCenter
               .Range(.Cells(8, 6), .Cells(8, 6)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(8, 6), .Cells(8, 6)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
               .Range(.Cells(8, 7), .Cells(8, 7)) = g_rst_Princi!cont
               .Range(.Cells(8, 7), .Cells(8, 7)).VerticalAlignment = xlCenter
               .Range(.Cells(8, 8), .Cells(8, 8)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(8, 8), .Cells(8, 8)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
               .Range(.Cells(8, 9), .Cells(8, 9)) = g_rst_Princi!cont
               .Range(.Cells(8, 9), .Cells(8, 9)).VerticalAlignment = xlCenter
               .Range(.Cells(8, 10), .Cells(8, 10)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(8, 10), .Cells(8, 10)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
               .Range(.Cells(8, 11), .Cells(8, 11)) = g_rst_Princi!cont
               .Range(.Cells(8, 11), .Cells(8, 11)).VerticalAlignment = xlCenter
               .Range(.Cells(8, 12), .Cells(8, 12)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(8, 12), .Cells(8, 12)).VerticalAlignment = xlCenter
            End If
            
            g_rst_Princi.MoveNext
         Loop
         
         'MICASITA
         g_str_Parame = ""
         g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
         g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD IN ('002','006','011')"
         g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
'         g_str_Parame = g_str_Parame + " UNION SELECT DISTINCT HIPCIE_CLAPRV, 0, 0 FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         Do While Not g_rst_Princi.EOF
            If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
               .Range(.Cells(10, 3), .Cells(10, 3)) = g_rst_Princi!cont
               .Range(.Cells(10, 3), .Cells(10, 3)).VerticalAlignment = xlCenter
               .Range(.Cells(10, 4), .Cells(10, 4)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(10, 4), .Cells(10, 4)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
               .Range(.Cells(10, 5), .Cells(10, 5)) = g_rst_Princi!cont
               .Range(.Cells(10, 5), .Cells(10, 5)).VerticalAlignment = xlCenter
               .Range(.Cells(10, 6), .Cells(10, 6)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(10, 6), .Cells(10, 6)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
               .Range(.Cells(10, 7), .Cells(10, 7)) = g_rst_Princi!cont
               .Range(.Cells(10, 7), .Cells(10, 7)).VerticalAlignment = xlCenter
               .Range(.Cells(10, 8), .Cells(10, 8)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(10, 8), .Cells(10, 8)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
               .Range(.Cells(10, 9), .Cells(10, 9)) = g_rst_Princi!cont
               .Range(.Cells(10, 9), .Cells(10, 9)).VerticalAlignment = xlCenter
               .Range(.Cells(10, 10), .Cells(10, 10)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(10, 10), .Cells(10, 10)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
               .Range(.Cells(10, 11), .Cells(10, 11)) = g_rst_Princi!cont
               .Range(.Cells(10, 11), .Cells(10, 11)).VerticalAlignment = xlCenter
               .Range(.Cells(10, 12), .Cells(10, 12)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(10, 12), .Cells(10, 12)).VerticalAlignment = xlCenter
            End If
            
            g_rst_Princi.MoveNext
         Loop
         
         'CME
         g_str_Parame = ""
         g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
         g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD = '003'"
         g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
'         g_str_Parame = g_str_Parame + " UNION SELECT DISTINCT HIPCIE_CLAPRV, 0, 0 FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         Do While Not g_rst_Princi.EOF
            If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
               .Range(.Cells(12, 3), .Cells(12, 3)) = g_rst_Princi!cont
               .Range(.Cells(12, 3), .Cells(12, 3)).VerticalAlignment = xlCenter
               .Range(.Cells(12, 4), .Cells(12, 4)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(12, 4), .Cells(12, 4)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
               .Range(.Cells(12, 5), .Cells(12, 5)) = g_rst_Princi!cont
               .Range(.Cells(12, 5), .Cells(12, 5)).VerticalAlignment = xlCenter
               .Range(.Cells(12, 6), .Cells(12, 6)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(12, 6), .Cells(12, 6)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
               .Range(.Cells(12, 7), .Cells(12, 7)) = g_rst_Princi!cont
               .Range(.Cells(12, 7), .Cells(12, 7)).VerticalAlignment = xlCenter
               .Range(.Cells(12, 8), .Cells(12, 8)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(12, 8), .Cells(12, 8)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
               .Range(.Cells(12, 9), .Cells(12, 9)) = g_rst_Princi!cont
               .Range(.Cells(12, 9), .Cells(12, 9)).VerticalAlignment = xlCenter
               .Range(.Cells(12, 10), .Cells(12, 10)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(12, 10), .Cells(12, 10)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
               .Range(.Cells(12, 11), .Cells(12, 11)) = g_rst_Princi!cont
               .Range(.Cells(12, 11), .Cells(12, 11)).VerticalAlignment = xlCenter
               .Range(.Cells(12, 12), .Cells(12, 12)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(12, 12), .Cells(12, 12)).VerticalAlignment = xlCenter
            End If
            
            g_rst_Princi.MoveNext
         Loop

         'MIVIVIENDA
         g_str_Parame = ""
         g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
         g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','023')"
         g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
         g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         Do While Not g_rst_Princi.EOF
            If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
               .Range(.Cells(14, 3), .Cells(14, 3)) = g_rst_Princi!cont
               .Range(.Cells(14, 3), .Cells(14, 3)).VerticalAlignment = xlCenter
               .Range(.Cells(14, 4), .Cells(14, 4)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(14, 4), .Cells(14, 4)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
               .Range(.Cells(14, 5), .Cells(14, 5)) = g_rst_Princi!cont
               .Range(.Cells(14, 5), .Cells(14, 5)).VerticalAlignment = xlCenter
               .Range(.Cells(14, 6), .Cells(14, 6)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(14, 6), .Cells(14, 6)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
               .Range(.Cells(14, 7), .Cells(14, 7)) = g_rst_Princi!cont
               .Range(.Cells(14, 7), .Cells(14, 7)).VerticalAlignment = xlCenter
               .Range(.Cells(14, 8), .Cells(14, 8)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(14, 8), .Cells(14, 8)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
               .Range(.Cells(14, 9), .Cells(14, 9)) = g_rst_Princi!cont
               .Range(.Cells(14, 9), .Cells(14, 9)).VerticalAlignment = xlCenter
               .Range(.Cells(14, 10), .Cells(14, 10)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(14, 10), .Cells(14, 10)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
               .Range(.Cells(14, 11), .Cells(14, 11)) = g_rst_Princi!cont
               .Range(.Cells(14, 11), .Cells(14, 11)).VerticalAlignment = xlCenter
               .Range(.Cells(14, 12), .Cells(14, 12)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(14, 12), .Cells(14, 12)).VerticalAlignment = xlCenter
            End If
                       
            g_rst_Princi.MoveNext
         Loop

         'MICASAMAS
         g_str_Parame = ""
         g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
         g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD = '019'"
         g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
         g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         Do While Not g_rst_Princi.EOF
            If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
               .Range(.Cells(16, 3), .Cells(16, 3)) = g_rst_Princi!cont
               .Range(.Cells(16, 3), .Cells(16, 3)).VerticalAlignment = xlCenter
               .Range(.Cells(16, 4), .Cells(16, 4)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(16, 4), .Cells(16, 4)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
               .Range(.Cells(16, 5), .Cells(16, 5)) = g_rst_Princi!cont
               .Range(.Cells(16, 5), .Cells(16, 5)).VerticalAlignment = xlCenter
               .Range(.Cells(16, 6), .Cells(16, 6)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(16, 6), .Cells(16, 6)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
               .Range(.Cells(16, 7), .Cells(16, 7)) = g_rst_Princi!cont
               .Range(.Cells(16, 7), .Cells(16, 7)).VerticalAlignment = xlCenter
               .Range(.Cells(16, 8), .Cells(16, 8)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(16, 8), .Cells(16, 8)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
               .Range(.Cells(16, 9), .Cells(16, 9)) = g_rst_Princi!cont
               .Range(.Cells(16, 9), .Cells(16, 9)).VerticalAlignment = xlCenter
               .Range(.Cells(16, 10), .Cells(16, 10)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(16, 10), .Cells(16, 10)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
               .Range(.Cells(16, 11), .Cells(16, 11)) = g_rst_Princi!cont
               .Range(.Cells(16, 11), .Cells(16, 11)).VerticalAlignment = xlCenter
               .Range(.Cells(16, 12), .Cells(16, 12)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(16, 12), .Cells(16, 12)).VerticalAlignment = xlCenter
            End If
                       
            g_rst_Princi.MoveNext
         Loop
         
         'MIVIVIENDA MAS
         g_str_Parame = ""
         g_str_Parame = g_str_Parame + "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
         g_str_Parame = g_str_Parame + "  FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame + " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD = '021'"
         g_str_Parame = g_str_Parame + " GROUP BY HIPCIE_CLAPRV"
         g_str_Parame = g_str_Parame + " ORDER BY HIPCIE_CLAPRV"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         Do While Not g_rst_Princi.EOF
            If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
               .Range(.Cells(18, 3), .Cells(18, 3)) = g_rst_Princi!cont
               .Range(.Cells(18, 3), .Cells(18, 3)).VerticalAlignment = xlCenter
               .Range(.Cells(18, 4), .Cells(18, 4)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(18, 4), .Cells(18, 4)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
               .Range(.Cells(18, 5), .Cells(18, 5)) = g_rst_Princi!cont
               .Range(.Cells(18, 5), .Cells(18, 5)).VerticalAlignment = xlCenter
               .Range(.Cells(18, 6), .Cells(18, 6)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(18, 6), .Cells(18, 6)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
               .Range(.Cells(18, 7), .Cells(18, 7)) = g_rst_Princi!cont
               .Range(.Cells(18, 7), .Cells(18, 7)).VerticalAlignment = xlCenter
               .Range(.Cells(18, 8), .Cells(18, 8)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(18, 8), .Cells(18, 8)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
               .Range(.Cells(18, 9), .Cells(18, 9)) = g_rst_Princi!cont
               .Range(.Cells(18, 9), .Cells(18, 9)).VerticalAlignment = xlCenter
               .Range(.Cells(18, 10), .Cells(18, 10)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(18, 10), .Cells(18, 10)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
               .Range(.Cells(18, 11), .Cells(18, 11)) = g_rst_Princi!cont
               .Range(.Cells(18, 11), .Cells(18, 11)).VerticalAlignment = xlCenter
               .Range(.Cells(18, 12), .Cells(18, 12)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(18, 12), .Cells(18, 12)).VerticalAlignment = xlCenter
            End If
            
            g_rst_Princi.MoveNext
         Loop

         'BBP
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT HIPCIE_CLAPRV, COUNT(*) AS CONT, ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_SALCAP+HIPCIE_SALCON, (HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM)),2) AS SALDO"
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCIE_PERANO = " & ipp_PerAno.Text & " AND HIPCIE_CODPRD = '022' "
         g_str_Parame = g_str_Parame & " GROUP BY HIPCIE_CLAPRV"
         g_str_Parame = g_str_Parame & " ORDER BY HIPCIE_CLAPRV"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         Do While Not g_rst_Princi.EOF
            If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
               .Range(.Cells(20, 3), .Cells(20, 3)) = g_rst_Princi!cont
               .Range(.Cells(20, 3), .Cells(20, 3)).VerticalAlignment = xlCenter
               .Range(.Cells(20, 4), .Cells(20, 4)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(20, 4), .Cells(20, 4)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
               .Range(.Cells(20, 5), .Cells(20, 5)) = g_rst_Princi!cont
               .Range(.Cells(20, 5), .Cells(20, 5)).VerticalAlignment = xlCenter
               .Range(.Cells(20, 6), .Cells(20, 6)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(20, 6), .Cells(20, 6)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 2 Then
               .Range(.Cells(20, 7), .Cells(20, 7)) = g_rst_Princi!cont
               .Range(.Cells(20, 7), .Cells(20, 7)).VerticalAlignment = xlCenter
               .Range(.Cells(20, 8), .Cells(20, 8)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(20, 8), .Cells(20, 8)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 3 Then
               .Range(.Cells(20, 9), .Cells(20, 9)) = g_rst_Princi!cont
               .Range(.Cells(20, 9), .Cells(20, 9)).VerticalAlignment = xlCenter
               .Range(.Cells(20, 10), .Cells(20, 10)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(20, 10), .Cells(20, 10)).VerticalAlignment = xlCenter
            End If
            
            If g_rst_Princi!HIPCIE_CLAPRV = 4 Then
               .Range(.Cells(20, 11), .Cells(20, 11)) = g_rst_Princi!cont
               .Range(.Cells(20, 11), .Cells(20, 11)).VerticalAlignment = xlCenter
               .Range(.Cells(20, 12), .Cells(20, 12)) = CLng(g_rst_Princi!SALDO)
               .Range(.Cells(20, 12), .Cells(20, 12)).VerticalAlignment = xlCenter
            End If
                       
            g_rst_Princi.MoveNext
         Loop

         .Cells(8, 13).Formula = "=C8+E8+G8+I8+K8"
         .Cells(9, 13).Formula = "=D8+F8+H8+J8+L8"
         .Cells(10, 13).Formula = "=C10+E10+G10+I10+K10"
         .Cells(11, 13).Formula = "=D10+F10+H10+J10+L10"
         .Cells(12, 13).Formula = "=C12+E12+G12+I12+K12"
         .Cells(13, 13).Formula = "=D12+F12+H12+J12+L12"
         .Cells(14, 13).Formula = "=C14+E14+G14+I14+K14"
         .Cells(15, 13).Formula = "=D14+F14+H14+J14+L14"
         .Cells(16, 13).Formula = "=C16+E16+G16+I16+K16"
         .Cells(17, 13).Formula = "=D16+F16+H16+J16+L16"
         .Cells(18, 13).Formula = "=C18+E18+G18+I18+K18"
         .Cells(19, 13).Formula = "=D18+F18+H18+J18+L18"
         .Cells(20, 13).Formula = "=C20+E20+G20+I20+K20"
         .Cells(21, 13).Formula = "=D20+F20+H20+J20+L20"
         
         .Cells(24, 3).Formula = "=C8+C10+C12+C14+C16+C18+C20"
         .Cells(24, 4).Formula = "=D8+D10+D12+D14+D16+D18+D20"
         .Cells(24, 5).Formula = "=E8+E10+E12+E14+E16+E18+E20"
         .Cells(24, 6).Formula = "=F8+F10+F12+F14+F16+F18+F20"
         .Cells(24, 7).Formula = "=G8+G10+G12+G14+G16+G18+G20"
         .Cells(24, 8).Formula = "=H8+H10+H12+H14+H16+H18+H20"
         .Cells(24, 9).Formula = "=I8+I10+I12+I14+I16+I18+I20"
         .Cells(24, 10).Formula = "=J8+J10+J12+J14+J16+J18+J20"
         .Cells(24, 11).Formula = "=K8+K10+K12+K14+K16+K18+K20"
         .Cells(24, 12).Formula = "=L8+L10+L12+L14+L16+L18+L20"
         
         .Cells(22, 13).Formula = "=M8+M10+M12+M14+M16+M18+M20"
         .Cells(23, 13).Formula = "=M9+M11+M13+M15+M17+M19+M21"
         
         .Cells(9, 3).Formula = "=C8/M8"
         .Cells(9, 4).Formula = "=D8/M9"
         .Cells(9, 5).Formula = "=E8/M8"
         .Cells(9, 6).Formula = "=F8/M9"
         .Cells(9, 7).Formula = "=G8/M8"
         .Cells(9, 8).Formula = "=H8/M9"
         .Cells(9, 9).Formula = "=I8/M8"
         .Cells(9, 10).Formula = "=J8/M9"
         .Cells(9, 11).Formula = "=K8/M8"
         .Cells(9, 12).Formula = "=L8/M9"
         
         .Cells(11, 3).Formula = "=C10/M10"
         .Cells(11, 4).Formula = "=D10/M11"
         .Cells(11, 5).Formula = "=E10/M10"
         .Cells(11, 6).Formula = "=F10/M11"
         .Cells(11, 7).Formula = "=G10/M10"
         .Cells(11, 8).Formula = "=H10/M11"
         .Cells(11, 9).Formula = "=I10/M10"
         .Cells(11, 10).Formula = "=J10/M11"
         .Cells(11, 11).Formula = "=K10/M10"
         .Cells(11, 12).Formula = "=L10/M11"
         
         .Cells(13, 3).Formula = "=C12/M12"
         .Cells(13, 4).Formula = "=D12/M13"
         .Cells(13, 5).Formula = "=E12/M12"
         .Cells(13, 6).Formula = "=F12/M13"
         .Cells(13, 7).Formula = "=G12/M12"
         .Cells(13, 8).Formula = "=H12/M13"
         .Cells(13, 9).Formula = "=I12/M12"
         .Cells(13, 10).Formula = "=J12/M13"
         .Cells(13, 11).Formula = "=K12/M12"
         .Cells(13, 12).Formula = "=L12/M13"
         
         .Cells(15, 3).Formula = "=C14/M14"
         .Cells(15, 4).Formula = "=D14/M15"
         .Cells(15, 5).Formula = "=E14/M14"
         .Cells(15, 6).Formula = "=F14/M15"
         .Cells(15, 7).Formula = "=G14/M14"
         .Cells(15, 8).Formula = "=H14/M15"
         .Cells(15, 9).Formula = "=I14/M14"
         .Cells(15, 10).Formula = "=J14/M15"
         .Cells(15, 11).Formula = "=K14/M14"
         .Cells(15, 12).Formula = "=L14/M15"
         
         .Cells(17, 3).Formula = "=C16/M16"
         .Cells(17, 4).Formula = "=D16/M17"
         .Cells(17, 5).Formula = "=E16/M16"
         .Cells(17, 6).Formula = "=F16/M17"
         .Cells(17, 7).Formula = "=G16/M16"
         .Cells(17, 8).Formula = "=H16/M17"
         .Cells(17, 9).Formula = "=I16/M16"
         .Cells(17, 10).Formula = "=J16/M17"
         .Cells(17, 11).Formula = "=K16/M16"
         .Cells(17, 12).Formula = "=L16/M17"

         .Cells(19, 3).Formula = "=C18/M18"
         .Cells(19, 4).Formula = "=D18/M19"
         .Cells(19, 5).Formula = "=E18/M18"
         .Cells(19, 6).Formula = "=F18/M19"
         .Cells(19, 7).Formula = "=G18/M18"
         .Cells(19, 8).Formula = "=H18/M19"
         .Cells(19, 9).Formula = "=I18/M18"
         .Cells(19, 10).Formula = "=J18/M19"
         .Cells(19, 11).Formula = "=K18/M18"
         .Cells(19, 12).Formula = "=L18/M19"
         
         .Cells(21, 3).Formula = "=C20/M20"
         .Cells(21, 4).Formula = "=D20/M21"
         .Cells(21, 5).Formula = "=E20/M20"
         .Cells(21, 6).Formula = "=F20/M21"
         .Cells(21, 7).Formula = "=G20/M20"
         .Cells(21, 8).Formula = "=H20/M21"
         .Cells(21, 9).Formula = "=I20/M20"
         .Cells(21, 10).Formula = "=J20/M21"
         .Cells(21, 11).Formula = "=K20/M20"
         .Cells(21, 12).Formula = "=L20/M21"
         
         .Range(.Cells(22, 3), .Cells(23, 3)).Merge
         .Range(.Cells(22, 4), .Cells(23, 4)).Merge
         .Range(.Cells(22, 5), .Cells(23, 5)).Merge
         .Range(.Cells(22, 6), .Cells(23, 6)).Merge
         .Range(.Cells(22, 7), .Cells(23, 7)).Merge
         .Range(.Cells(22, 8), .Cells(23, 8)).Merge
         .Range(.Cells(22, 9), .Cells(23, 9)).Merge
         .Range(.Cells(22, 10), .Cells(23, 10)).Merge
         .Range(.Cells(22, 11), .Cells(23, 11)).Merge
         .Range(.Cells(22, 12), .Cells(23, 12)).Merge
         .Range(.Cells(22, 3), .Cells(23, 12)).VerticalAlignment = xlCenter
         
         .Cells(22, 3).Formula = "=C24/M22"
         .Cells(22, 4).Formula = "=D24/M23"
         .Cells(22, 5).Formula = "=E24/M22"
         .Cells(22, 6).Formula = "=F24/M23"
         .Cells(22, 7).Formula = "=G24/M22"
         .Cells(22, 8).Formula = "=H24/M23"
         .Cells(22, 9).Formula = "=I24/M22"
         .Cells(22, 10).Formula = "=J24/M23"
         .Cells(22, 11).Formula = "=K24/M22"
         .Cells(22, 12).Formula = "=L24/M23"
         
         .Range(.Cells(22, 2), .Cells(23, 13)).Font.Bold = True
         .Range(.Cells(24, 2), .Cells(24, 13)).Font.Bold = True
         
         .Range(.Cells(9, 3), .Cells(9, 12)).NumberFormat = "0.00%"
         .Range(.Cells(11, 3), .Cells(11, 12)).NumberFormat = "0.00%"
         .Range(.Cells(13, 3), .Cells(13, 12)).NumberFormat = "0.00%"
         .Range(.Cells(15, 3), .Cells(15, 12)).NumberFormat = "0.00%"
         .Range(.Cells(17, 3), .Cells(17, 12)).NumberFormat = "0.00%"
         .Range(.Cells(19, 3), .Cells(19, 12)).NumberFormat = "0.00%"
         .Range(.Cells(21, 3), .Cells(21, 12)).NumberFormat = "0.00%"
         .Range(.Cells(22, 3), .Cells(22, 12)).NumberFormat = "0.00%"
         
         For r_int_Contad = 8 To 24
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(r_int_Contad, 3), .Cells(r_int_Contad, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next
         
         
      Else
         .Cells(1, 1) = "CONSOLIDADO DE CLASIFICION DE CARTERA A " & UCase(Trim(cmb_PerMes.Text)) & " DEL " & CStr(ipp_PerAno.Text) & " EN SOLES "
         .Range(.Cells(1, 1), .Cells(1, 37)).Merge
         .Range("A1:Y1").HorizontalAlignment = xlHAlignCenter
         
         'Primera Linea
         r_int_NroFil = 3
         .Cells(r_int_NroFil, 1) = "CLASIFICACION"
         .Cells(r_int_NroFil, 2) = "ENERO"
         .Cells(r_int_NroFil, 5) = "FEBRERO"
         .Cells(r_int_NroFil, 8) = "MARZO"
         .Cells(r_int_NroFil, 11) = "ABRIL"
         .Cells(r_int_NroFil, 14) = "MAYO"
         .Cells(r_int_NroFil, 17) = "JUNIO"
         .Cells(r_int_NroFil, 20) = "JULIO"
         .Cells(r_int_NroFil, 23) = "AGOSTO"
         .Cells(r_int_NroFil, 26) = "SETIEMBRE"
         .Cells(r_int_NroFil, 29) = "OCTUBRE"
         .Cells(r_int_NroFil, 32) = "NOVIEMBRE"
         .Cells(r_int_NroFil, 35) = "DICIEMBRE"
         .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil, 4)).Merge
         .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 7)).Merge
         .Range(.Cells(r_int_NroFil, 8), .Cells(r_int_NroFil, 10)).Merge
         .Range(.Cells(r_int_NroFil, 11), .Cells(r_int_NroFil, 13)).Merge
         .Range(.Cells(r_int_NroFil, 14), .Cells(r_int_NroFil, 16)).Merge
         .Range(.Cells(r_int_NroFil, 17), .Cells(r_int_NroFil, 19)).Merge
         .Range(.Cells(r_int_NroFil, 20), .Cells(r_int_NroFil, 22)).Merge
         .Range(.Cells(r_int_NroFil, 23), .Cells(r_int_NroFil, 25)).Merge
         .Range(.Cells(r_int_NroFil, 26), .Cells(r_int_NroFil, 28)).Merge
         .Range(.Cells(r_int_NroFil, 29), .Cells(r_int_NroFil, 31)).Merge
         .Range(.Cells(r_int_NroFil, 32), .Cells(r_int_NroFil, 34)).Merge
         .Range(.Cells(r_int_NroFil, 35), .Cells(r_int_NroFil, 37)).Merge
         
         'Segunda Linea
         r_int_NroFil = r_int_NroFil + 1
         .Columns("A").ColumnWidth = 15
         .Columns("B").ColumnWidth = 9:    .Cells(r_int_NroFil, 2) = "NUMERO":          .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignCenter
         .Columns("C").ColumnWidth = 13:   .Cells(r_int_NroFil, 3) = "MONTO S/.":       .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignRight
         .Columns("D").ColumnWidth = 13:   .Cells(r_int_NroFil, 4) = "PROVISION S/.":   .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignRight
         .Columns("E").ColumnWidth = 9:    .Cells(r_int_NroFil, 5) = "NUMERO":          .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignCenter
         .Columns("F").ColumnWidth = 13:   .Cells(r_int_NroFil, 6) = "MONTO":           .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignRight
         .Columns("G").ColumnWidth = 13:   .Cells(r_int_NroFil, 7) = "PROVISION S/.":   .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignRight
         .Columns("H").ColumnWidth = 9:    .Cells(r_int_NroFil, 8) = "NUMERO":          .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignCenter
         .Columns("I").ColumnWidth = 13:   .Cells(r_int_NroFil, 9) = "MONTO":           .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignRight
         .Columns("J").ColumnWidth = 13:   .Cells(r_int_NroFil, 10) = "PROVISION S/.":  .Cells(r_int_NroFil, 10).HorizontalAlignment = xlHAlignRight
         .Columns("K").ColumnWidth = 9:    .Cells(r_int_NroFil, 11) = "NUMERO":         .Cells(r_int_NroFil, 11).HorizontalAlignment = xlHAlignCenter
         .Columns("L").ColumnWidth = 13:   .Cells(r_int_NroFil, 12) = "MONTO":          .Cells(r_int_NroFil, 12).HorizontalAlignment = xlHAlignRight
         .Columns("M").ColumnWidth = 13:   .Cells(r_int_NroFil, 13) = "PROVISION S/.":  .Cells(r_int_NroFil, 13).HorizontalAlignment = xlHAlignRight
         .Columns("N").ColumnWidth = 9:    .Cells(r_int_NroFil, 14) = "NUMERO":         .Cells(r_int_NroFil, 14).HorizontalAlignment = xlHAlignCenter
         .Columns("O").ColumnWidth = 13:   .Cells(r_int_NroFil, 15) = "MONTO":          .Cells(r_int_NroFil, 15).HorizontalAlignment = xlHAlignRight
         .Columns("P").ColumnWidth = 13:   .Cells(r_int_NroFil, 16) = "PROVISION S/.":  .Cells(r_int_NroFil, 16).HorizontalAlignment = xlHAlignRight
         .Columns("Q").ColumnWidth = 9:    .Cells(r_int_NroFil, 17) = "NUMERO":         .Cells(r_int_NroFil, 17).HorizontalAlignment = xlHAlignCenter
         .Columns("R").ColumnWidth = 13:   .Cells(r_int_NroFil, 18) = "MONTO":          .Cells(r_int_NroFil, 18).HorizontalAlignment = xlHAlignRight
         .Columns("S").ColumnWidth = 13:   .Cells(r_int_NroFil, 19) = "PROVISION S/.":  .Cells(r_int_NroFil, 19).HorizontalAlignment = xlHAlignRight
         .Columns("T").ColumnWidth = 9:    .Cells(r_int_NroFil, 20) = "NUMERO":         .Cells(r_int_NroFil, 20).HorizontalAlignment = xlHAlignCenter
         .Columns("U").ColumnWidth = 13:   .Cells(r_int_NroFil, 21) = "MONTO":          .Cells(r_int_NroFil, 21).HorizontalAlignment = xlHAlignRight
         .Columns("V").ColumnWidth = 13:   .Cells(r_int_NroFil, 22) = "PROVISION S/.":  .Cells(r_int_NroFil, 22).HorizontalAlignment = xlHAlignRight
         .Columns("W").ColumnWidth = 9:    .Cells(r_int_NroFil, 23) = "NUMERO":         .Cells(r_int_NroFil, 23).HorizontalAlignment = xlHAlignCenter
         .Columns("X").ColumnWidth = 13:   .Cells(r_int_NroFil, 24) = "MONTO":          .Cells(r_int_NroFil, 24).HorizontalAlignment = xlHAlignRight
         .Columns("Y").ColumnWidth = 13:   .Cells(r_int_NroFil, 25) = "PROVISION S/.":  .Cells(r_int_NroFil, 25).HorizontalAlignment = xlHAlignRight
         .Columns("Z").ColumnWidth = 9:    .Cells(r_int_NroFil, 26) = "NUMERO":         .Cells(r_int_NroFil, 26).HorizontalAlignment = xlHAlignCenter
         .Columns("AA").ColumnWidth = 13:  .Cells(r_int_NroFil, 27) = "MONTO":          .Cells(r_int_NroFil, 27).HorizontalAlignment = xlHAlignRight
         .Columns("AB").ColumnWidth = 13:  .Cells(r_int_NroFil, 28) = "PROVISION S/.":  .Cells(r_int_NroFil, 28).HorizontalAlignment = xlHAlignRight
         .Columns("AC").ColumnWidth = 9:   .Cells(r_int_NroFil, 29) = "NUMERO":         .Cells(r_int_NroFil, 29).HorizontalAlignment = xlHAlignCenter
         .Columns("AD").ColumnWidth = 13:  .Cells(r_int_NroFil, 30) = "MONTO":          .Cells(r_int_NroFil, 30).HorizontalAlignment = xlHAlignRight
         .Columns("AE").ColumnWidth = 13:  .Cells(r_int_NroFil, 31) = "PROVISION S/.":  .Cells(r_int_NroFil, 31).HorizontalAlignment = xlHAlignRight
         .Columns("AF").ColumnWidth = 9:   .Cells(r_int_NroFil, 32) = "NUMERO":         .Cells(r_int_NroFil, 32).HorizontalAlignment = xlHAlignCenter
         .Columns("AG").ColumnWidth = 13:  .Cells(r_int_NroFil, 33) = "MONTO":          .Cells(r_int_NroFil, 33).HorizontalAlignment = xlHAlignRight
         .Columns("AH").ColumnWidth = 13:  .Cells(r_int_NroFil, 34) = "PROVISION S/.":  .Cells(r_int_NroFil, 34).HorizontalAlignment = xlHAlignRight
         .Columns("AI").ColumnWidth = 9:   .Cells(r_int_NroFil, 35) = "NUMERO":         .Cells(r_int_NroFil, 35).HorizontalAlignment = xlHAlignCenter
         .Columns("AJ").ColumnWidth = 13:  .Cells(r_int_NroFil, 36) = "MONTO":          .Cells(r_int_NroFil, 36).HorizontalAlignment = xlHAlignRight
         .Columns("AK").ColumnWidth = 13:  .Cells(r_int_NroFil, 37) = "PROVISION S/.":  .Cells(r_int_NroFil, 37).HorizontalAlignment = xlHAlignRight
         
         'Combina celdas de primer linea
         .Range("B3:D4").HorizontalAlignment = xlHAlignCenter
         .Range("E3:G4").HorizontalAlignment = xlHAlignCenter
         .Range("H3:J4").HorizontalAlignment = xlHAlignCenter
         .Range("K3:M4").HorizontalAlignment = xlHAlignCenter
         .Range("N3:P4").HorizontalAlignment = xlHAlignCenter
         .Range("Q3:S4").HorizontalAlignment = xlHAlignCenter
         .Range("T3:V4").HorizontalAlignment = xlHAlignCenter
         .Range("W3:Y4").HorizontalAlignment = xlHAlignCenter
         .Range("Z3:AB4").HorizontalAlignment = xlHAlignCenter
         .Range("AC3:AE4").HorizontalAlignment = xlHAlignCenter
         .Range("AF3:AH4").HorizontalAlignment = xlHAlignCenter
         .Range("AI3:AK4").HorizontalAlignment = xlHAlignCenter
         
         'Formatea titulo
         .Range(.Cells(1, 1), .Cells(r_int_NroFil, 37)).Font.Name = "Calibri"
         .Range(.Cells(1, 1), .Cells(r_int_NroFil, 37)).Font.Size = 11
         .Range(.Cells(1, 1), .Cells(r_int_NroFil, 37)).Font.Bold = True
         
         'Exporta filas
         If CInt(tab_Clasif.Tab) = 0 Then
            For r_int_Contad = 5 To 12
               .Cells(r_int_Contad, 1) = grd_LisCab.TextMatrix(r_int_Contad - 3, 1)
               .Cells(r_int_Contad, 2) = grd_LisCab.TextMatrix(r_int_Contad - 3, 2)
               .Cells(r_int_Contad, 3) = grd_LisCab.TextMatrix(r_int_Contad - 3, 3)
               .Cells(r_int_Contad, 4) = grd_LisCab.TextMatrix(r_int_Contad - 3, 4)
               .Cells(r_int_Contad, 5) = grd_LisCab.TextMatrix(r_int_Contad - 3, 5)
               .Cells(r_int_Contad, 6) = grd_LisCab.TextMatrix(r_int_Contad - 3, 6)
               .Cells(r_int_Contad, 7) = grd_LisCab.TextMatrix(r_int_Contad - 3, 7)
               .Cells(r_int_Contad, 8) = grd_LisCab.TextMatrix(r_int_Contad - 3, 8)
               .Cells(r_int_Contad, 9) = grd_LisCab.TextMatrix(r_int_Contad - 3, 9)
               .Cells(r_int_Contad, 10) = grd_LisCab.TextMatrix(r_int_Contad - 3, 10)
               .Cells(r_int_Contad, 11) = grd_LisCab.TextMatrix(r_int_Contad - 3, 11)
               .Cells(r_int_Contad, 12) = grd_LisCab.TextMatrix(r_int_Contad - 3, 12)
               .Cells(r_int_Contad, 13) = grd_LisCab.TextMatrix(r_int_Contad - 3, 13)
               .Cells(r_int_Contad, 14) = grd_LisCab.TextMatrix(r_int_Contad - 3, 14)
               .Cells(r_int_Contad, 15) = grd_LisCab.TextMatrix(r_int_Contad - 3, 15)
               .Cells(r_int_Contad, 16) = grd_LisCab.TextMatrix(r_int_Contad - 3, 16)
               .Cells(r_int_Contad, 17) = grd_LisCab.TextMatrix(r_int_Contad - 3, 17)
               .Cells(r_int_Contad, 18) = grd_LisCab.TextMatrix(r_int_Contad - 3, 18)
               .Cells(r_int_Contad, 19) = grd_LisCab.TextMatrix(r_int_Contad - 3, 19)
               .Cells(r_int_Contad, 20) = grd_LisCab.TextMatrix(r_int_Contad - 3, 20)
               .Cells(r_int_Contad, 21) = grd_LisCab.TextMatrix(r_int_Contad - 3, 21)
               .Cells(r_int_Contad, 22) = grd_LisCab.TextMatrix(r_int_Contad - 3, 22)
               .Cells(r_int_Contad, 23) = grd_LisCab.TextMatrix(r_int_Contad - 3, 23)
               .Cells(r_int_Contad, 24) = grd_LisCab.TextMatrix(r_int_Contad - 3, 24)
               .Cells(r_int_Contad, 25) = grd_LisCab.TextMatrix(r_int_Contad - 3, 25)
               .Cells(r_int_Contad, 26) = grd_LisCab.TextMatrix(r_int_Contad - 3, 26)
               .Cells(r_int_Contad, 27) = grd_LisCab.TextMatrix(r_int_Contad - 3, 27)
               .Cells(r_int_Contad, 28) = grd_LisCab.TextMatrix(r_int_Contad - 3, 28)
               .Cells(r_int_Contad, 29) = grd_LisCab.TextMatrix(r_int_Contad - 3, 29)
               .Cells(r_int_Contad, 30) = grd_LisCab.TextMatrix(r_int_Contad - 3, 30)
               .Cells(r_int_Contad, 31) = grd_LisCab.TextMatrix(r_int_Contad - 3, 31)
               .Cells(r_int_Contad, 32) = grd_LisCab.TextMatrix(r_int_Contad - 3, 32)
               .Cells(r_int_Contad, 33) = grd_LisCab.TextMatrix(r_int_Contad - 3, 33)
               .Cells(r_int_Contad, 34) = grd_LisCab.TextMatrix(r_int_Contad - 3, 34)
               .Cells(r_int_Contad, 35) = grd_LisCab.TextMatrix(r_int_Contad - 3, 35)
               .Cells(r_int_Contad, 36) = grd_LisCab.TextMatrix(r_int_Contad - 3, 36)
               .Cells(r_int_Contad, 37) = grd_LisCab.TextMatrix(r_int_Contad - 3, 37)
            Next
         Else
            For r_int_Contad = 5 To 23
               .Cells(r_int_Contad, 1) = grd_LisDet.TextMatrix(r_int_Contad - 3, 1)
               .Cells(r_int_Contad, 2) = grd_LisDet.TextMatrix(r_int_Contad - 3, 2)
               .Cells(r_int_Contad, 3) = grd_LisDet.TextMatrix(r_int_Contad - 3, 3)
               .Cells(r_int_Contad, 4) = grd_LisDet.TextMatrix(r_int_Contad - 3, 4)
               .Cells(r_int_Contad, 5) = grd_LisDet.TextMatrix(r_int_Contad - 3, 5)
               .Cells(r_int_Contad, 6) = grd_LisDet.TextMatrix(r_int_Contad - 3, 6)
               .Cells(r_int_Contad, 7) = grd_LisDet.TextMatrix(r_int_Contad - 3, 7)
               .Cells(r_int_Contad, 8) = grd_LisDet.TextMatrix(r_int_Contad - 3, 8)
               .Cells(r_int_Contad, 9) = grd_LisDet.TextMatrix(r_int_Contad - 3, 9)
               .Cells(r_int_Contad, 10) = grd_LisDet.TextMatrix(r_int_Contad - 3, 10)
               .Cells(r_int_Contad, 11) = grd_LisDet.TextMatrix(r_int_Contad - 3, 11)
               .Cells(r_int_Contad, 12) = grd_LisDet.TextMatrix(r_int_Contad - 3, 12)
               .Cells(r_int_Contad, 13) = grd_LisDet.TextMatrix(r_int_Contad - 3, 13)
               .Cells(r_int_Contad, 14) = grd_LisDet.TextMatrix(r_int_Contad - 3, 14)
               .Cells(r_int_Contad, 15) = grd_LisDet.TextMatrix(r_int_Contad - 3, 15)
               .Cells(r_int_Contad, 16) = grd_LisDet.TextMatrix(r_int_Contad - 3, 16)
               .Cells(r_int_Contad, 17) = grd_LisDet.TextMatrix(r_int_Contad - 3, 17)
               .Cells(r_int_Contad, 18) = grd_LisDet.TextMatrix(r_int_Contad - 3, 18)
               .Cells(r_int_Contad, 19) = grd_LisDet.TextMatrix(r_int_Contad - 3, 19)
               .Cells(r_int_Contad, 20) = grd_LisDet.TextMatrix(r_int_Contad - 3, 20)
               .Cells(r_int_Contad, 21) = grd_LisDet.TextMatrix(r_int_Contad - 3, 21)
               .Cells(r_int_Contad, 22) = grd_LisDet.TextMatrix(r_int_Contad - 3, 22)
               .Cells(r_int_Contad, 23) = grd_LisDet.TextMatrix(r_int_Contad - 3, 23)
               .Cells(r_int_Contad, 24) = grd_LisDet.TextMatrix(r_int_Contad - 3, 24)
               .Cells(r_int_Contad, 25) = grd_LisDet.TextMatrix(r_int_Contad - 3, 25)
               .Cells(r_int_Contad, 26) = grd_LisDet.TextMatrix(r_int_Contad - 3, 26)
               .Cells(r_int_Contad, 27) = grd_LisDet.TextMatrix(r_int_Contad - 3, 27)
               .Cells(r_int_Contad, 28) = grd_LisDet.TextMatrix(r_int_Contad - 3, 28)
               .Cells(r_int_Contad, 29) = grd_LisDet.TextMatrix(r_int_Contad - 3, 29)
               .Cells(r_int_Contad, 30) = grd_LisDet.TextMatrix(r_int_Contad - 3, 30)
               .Cells(r_int_Contad, 31) = grd_LisDet.TextMatrix(r_int_Contad - 3, 31)
               .Cells(r_int_Contad, 32) = grd_LisDet.TextMatrix(r_int_Contad - 3, 32)
               .Cells(r_int_Contad, 33) = grd_LisDet.TextMatrix(r_int_Contad - 3, 33)
               .Cells(r_int_Contad, 34) = grd_LisDet.TextMatrix(r_int_Contad - 3, 34)
               .Cells(r_int_Contad, 35) = grd_LisDet.TextMatrix(r_int_Contad - 3, 35)
               .Cells(r_int_Contad, 36) = grd_LisDet.TextMatrix(r_int_Contad - 3, 36)
               .Cells(r_int_Contad, 37) = grd_LisDet.TextMatrix(r_int_Contad - 3, 37)
            Next
         End If
      End If
   End With
      
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
