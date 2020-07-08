VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_25 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   4290
   ClientTop       =   4665
   ClientWidth     =   14010
   Icon            =   "GesCtb_frm_851.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   14010
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel10 
      Height          =   675
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   13950
      _Version        =   65536
      _ExtentX        =   24606
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
         TabIndex        =   3
         Top             =   180
         Width           =   8925
         _Version        =   65536
         _ExtentX        =   15743
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Reporte de Consolidado de Clasificaciones de Cartera para Provisiones - Detalle"
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
         Picture         =   "GesCtb_frm_851.frx":000C
         Top             =   90
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   675
      Left            =   30
      TabIndex        =   4
      Top             =   750
      Width           =   13950
      _Version        =   65536
      _ExtentX        =   24606
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
      Font3D          =   2
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   45
         Picture         =   "GesCtb_frm_851.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Exportar a Excel"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   13320
         Picture         =   "GesCtb_frm_851.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   615
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   5820
      Left            =   30
      TabIndex        =   5
      Top             =   1440
      Width           =   13935
      _Version        =   65536
      _ExtentX        =   24580
      _ExtentY        =   10266
      _StockProps     =   15
      Caption         =   "SSPanel1"
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
         Height          =   525
         Left            =   90
         TabIndex        =   6
         Top             =   90
         Width           =   13755
         _Version        =   65536
         _ExtentX        =   24262
         _ExtentY        =   926
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
         Begin VB.Label Label5 
            Caption         =   "Clasificacion:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   150
            Width           =   1230
         End
         Begin VB.Label lblconcepto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1365
            TabIndex        =   7
            Top             =   120
            Width           =   12285
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5085
         Left            =   90
         TabIndex        =   9
         Top             =   660
         Width           =   13740
         _Version        =   65536
         _ExtentX        =   24236
         _ExtentY        =   8969
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisCla 
            Height          =   4965
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   13575
            _ExtentX        =   23945
            _ExtentY        =   8758
            _Version        =   393216
            Rows            =   18
            Cols            =   38
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
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
Attribute VB_Name = "frm_RptCtb_25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
   If moddat_g_int_TipCli = 1 Then
      Call fs_Buscar_Clasificacion_Cab
   Else
      Call fs_Buscar_Clasificacion_Det
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   lblConcepto.Caption = moddat_g_str_NomPrd
   
   If moddat_g_int_TipCli = 1 Then
      grd_LisCla.Rows = 12
   Else
      grd_LisCla.Rows = 18
   End If
   
   'Clasificacion
   grd_LisCla.ColWidth(0) = 0       ' CODIGO DE CLASIFICACION
   grd_LisCla.ColWidth(1) = 2000    ' DESCRIPCION DE CLASIFICACION
   grd_LisCla.ColWidth(2) = 900     ' NUMERO MES 01
   grd_LisCla.ColWidth(3) = 1300    ' MONTO MES 01
   grd_LisCla.ColWidth(4) = 1300    ' MONTO MES 01
   grd_LisCla.ColWidth(5) = 900     ' NUMERO MES 02
   grd_LisCla.ColWidth(6) = 1300    ' MONTO MES 02
   grd_LisCla.ColWidth(7) = 1300    ' MONTO MES 02
   grd_LisCla.ColWidth(8) = 900     ' NUMERO MES 03
   grd_LisCla.ColWidth(9) = 1300    ' MONTO MES 03
   grd_LisCla.ColWidth(10) = 1300    ' MONTO MES 03
   grd_LisCla.ColWidth(11) = 900     ' NUMERO MES 04
   grd_LisCla.ColWidth(12) = 1300    ' MONTO MES 04
   grd_LisCla.ColWidth(13) = 1300    ' MONTO MES 04
   grd_LisCla.ColWidth(14) = 900     ' NUMERO MES 05
   grd_LisCla.ColWidth(15) = 1300    ' MONTO MES 05
   grd_LisCla.ColWidth(16) = 1300    ' MONTO MES 05
   grd_LisCla.ColWidth(17) = 900     ' NUMERO MES 06
   grd_LisCla.ColWidth(18) = 1300    ' MONTO MES 06
   grd_LisCla.ColWidth(19) = 1300    ' MONTO MES 06
   grd_LisCla.ColWidth(20) = 900     ' NUMERO MES 07
   grd_LisCla.ColWidth(21) = 1300    ' MONTO MES 07
   grd_LisCla.ColWidth(22) = 1300    ' MONTO MES 07
   grd_LisCla.ColWidth(23) = 900     ' NUMERO MES 08
   grd_LisCla.ColWidth(24) = 1300    ' MONTO MES 08
   grd_LisCla.ColWidth(25) = 1300    ' MONTO MES 08
   grd_LisCla.ColWidth(26) = 900     ' NUMERO MES 09
   grd_LisCla.ColWidth(27) = 1300    ' MONTO MES 09
   grd_LisCla.ColWidth(28) = 1300    ' MONTO MES 09
   grd_LisCla.ColWidth(29) = 900     ' NUMERO MES 10
   grd_LisCla.ColWidth(30) = 1300    ' MONTO MES 10
   grd_LisCla.ColWidth(31) = 1300    ' MONTO MES 10
   grd_LisCla.ColWidth(32) = 900     ' NUMERO MES 11
   grd_LisCla.ColWidth(33) = 1300    ' MONTO MES 11
   grd_LisCla.ColWidth(34) = 1300    ' MONTO MES 11
   grd_LisCla.ColWidth(35) = 900     ' NUMERO MES 12
   grd_LisCla.ColWidth(36) = 1300    ' MONTO MES 12
   grd_LisCla.ColWidth(37) = 1300    ' MONTO MES 12
   grd_LisCla.ColAlignment(0) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(1) = flexAlignLeftCenter
   grd_LisCla.ColAlignment(2) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(3) = flexAlignRightCenter
   grd_LisCla.ColAlignment(4) = flexAlignRightCenter
   grd_LisCla.ColAlignment(5) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(6) = flexAlignRightCenter
   grd_LisCla.ColAlignment(7) = flexAlignRightCenter
   grd_LisCla.ColAlignment(8) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(9) = flexAlignRightCenter
   grd_LisCla.ColAlignment(10) = flexAlignRightCenter
   grd_LisCla.ColAlignment(11) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(12) = flexAlignRightCenter
   grd_LisCla.ColAlignment(13) = flexAlignRightCenter
   grd_LisCla.ColAlignment(14) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(15) = flexAlignRightCenter
   grd_LisCla.ColAlignment(16) = flexAlignRightCenter
   grd_LisCla.ColAlignment(17) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(18) = flexAlignRightCenter
   grd_LisCla.ColAlignment(19) = flexAlignRightCenter
   grd_LisCla.ColAlignment(20) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(21) = flexAlignRightCenter
   grd_LisCla.ColAlignment(22) = flexAlignRightCenter
   grd_LisCla.ColAlignment(23) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(24) = flexAlignRightCenter
   grd_LisCla.ColAlignment(25) = flexAlignRightCenter
   grd_LisCla.ColAlignment(26) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(27) = flexAlignRightCenter
   grd_LisCla.ColAlignment(28) = flexAlignRightCenter
   grd_LisCla.ColAlignment(29) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(30) = flexAlignRightCenter
   grd_LisCla.ColAlignment(31) = flexAlignRightCenter
   grd_LisCla.ColAlignment(32) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(33) = flexAlignRightCenter
   grd_LisCla.ColAlignment(34) = flexAlignRightCenter
   grd_LisCla.ColAlignment(35) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(36) = flexAlignRightCenter
   grd_LisCla.ColAlignment(37) = flexAlignRightCenter
   Call gs_LimpiaGrid(grd_LisCla)
End Sub

Private Sub fs_Buscar_Clasificacion_Cab()
Dim r_int_Contad     As Integer
Dim r_int_NumCol     As Integer
Dim r_int_NumFil     As Integer
Dim r_int_TotNum     As Integer
Dim r_dbl_TotMto     As Double
Dim r_dbl_TotPrv     As Double

   grd_LisCla.Redraw = False
   Call gs_LimpiaGrid(grd_LisCla)
   
   'Fila 0
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Row = 0:   grd_LisCla.Text = ""
   grd_LisCla.Col = 1:   grd_LisCla.Text = "PRODUCTOS"
   grd_LisCla.Col = 2:   grd_LisCla.Text = "ENERO"
   grd_LisCla.Col = 3:   grd_LisCla.Text = "ENERO"
   grd_LisCla.Col = 4:   grd_LisCla.Text = "ENERO"
   grd_LisCla.Col = 5:   grd_LisCla.Text = "FEBRERO"
   grd_LisCla.Col = 6:   grd_LisCla.Text = "FEBRERO"
   grd_LisCla.Col = 7:   grd_LisCla.Text = "FEBRERO"
   grd_LisCla.Col = 8:   grd_LisCla.Text = "MARZO"
   grd_LisCla.Col = 9:   grd_LisCla.Text = "MARZO"
   grd_LisCla.Col = 10:  grd_LisCla.Text = "MARZO"
   grd_LisCla.Col = 11:  grd_LisCla.Text = "ABRIL"
   grd_LisCla.Col = 12:  grd_LisCla.Text = "ABRIL"
   grd_LisCla.Col = 13:  grd_LisCla.Text = "ABRIL"
   grd_LisCla.Col = 14:  grd_LisCla.Text = "MAYO"
   grd_LisCla.Col = 15:  grd_LisCla.Text = "MAYO"
   grd_LisCla.Col = 16:  grd_LisCla.Text = "MAYO"
   grd_LisCla.Col = 17:  grd_LisCla.Text = "JUNIO"
   grd_LisCla.Col = 18:  grd_LisCla.Text = "JUNIO"
   grd_LisCla.Col = 19:  grd_LisCla.Text = "JUNIO"
   grd_LisCla.Col = 20:  grd_LisCla.Text = "JULIO"
   grd_LisCla.Col = 21:  grd_LisCla.Text = "JULIO"
   grd_LisCla.Col = 22:  grd_LisCla.Text = "JULIO"
   grd_LisCla.Col = 23:  grd_LisCla.Text = "AGOSTO"
   grd_LisCla.Col = 24:  grd_LisCla.Text = "AGOSTO"
   grd_LisCla.Col = 25:  grd_LisCla.Text = "AGOSTO"
   grd_LisCla.Col = 26:  grd_LisCla.Text = "SETIEMBRE"
   grd_LisCla.Col = 27:  grd_LisCla.Text = "SETIEMBRE"
   grd_LisCla.Col = 28:  grd_LisCla.Text = "SETIEMBRE"
   grd_LisCla.Col = 29:  grd_LisCla.Text = "OCTUBRE"
   grd_LisCla.Col = 30:  grd_LisCla.Text = "OCTUBRE"
   grd_LisCla.Col = 31:  grd_LisCla.Text = "OCTUBRE"
   grd_LisCla.Col = 32:  grd_LisCla.Text = "NOVIEMBRE"
   grd_LisCla.Col = 33:  grd_LisCla.Text = "NOVIEMBRE"
   grd_LisCla.Col = 34:  grd_LisCla.Text = "NOVIEMBRE"
   grd_LisCla.Col = 35:  grd_LisCla.Text = "DICIEMBRE"
   grd_LisCla.Col = 36:  grd_LisCla.Text = "DICIEMBRE"
   grd_LisCla.Col = 37:  grd_LisCla.Text = "DICIEMBRE"
   
   'Fila 1
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 1:   grd_LisCla.Text = "PRODUCTOS"
   grd_LisCla.Col = 2:   grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 3:   grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 4:   grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 5:   grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 6:   grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 7:   grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 8:   grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 9:   grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 10:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 11:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 12:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 13:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 14:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 15:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 16:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 17:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 18:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 19:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 20:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 21:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 22:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 23:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 24:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 25:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 26:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 27:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 28:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 29:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 30:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 31:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 32:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 33:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 34:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 35:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 36:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 37:  grd_LisCla.Text = "PROVISION S/."
   
   'Fila 2
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "CME"
   
   'Fila 3
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "CRC-PBP"
   
   'Fila 4
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "MICASITA"
   
   'Fila 5
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "MIVIVIENDA"
   
   'Fila 7
   grd_LisCla.Rows = grd_LisCla.Rows + 2
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "TOTALES"
   
   With grd_LisCla
      .MergeCells = flexMergeFree
      .MergeCol(1) = True
      .MergeRow(0) = True
      .FixedCols = 2
      .FixedRows = 2
   End With
   
   For r_int_Contad = 1 To moddat_g_int_EdaMes
      'Prepara SP que trae consolidado mensual por clasificacion
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT PRODUCTO AS PRODUCTO, SUM(NUMERO) AS NUMERO, SUM(MONTO_TOTAL) AS MONTO_TOTAL, SUM(PROVISIONES) AS PROVISIONES "
      g_str_Parame = g_str_Parame & "  FROM (SELECT CASE WHEN HIPCIE_CODPRD='001' THEN '2' "  'CRC
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='002' THEN '3' "  'MIC
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='003' THEN '1' "  'CME
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='004' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='006' THEN '3' "  'MIC
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='007' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='009' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='010' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='011' THEN '3' "  'MIC
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='012' THEN '3' "  'MIC
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='013' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='014' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='015' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='016' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='017' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='018' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='019' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='021' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='022' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='023' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "               END AS PRODUCTO, "
      g_str_Parame = g_str_Parame & "               COUNT(*) AS NUMERO, "
      g_str_Parame = g_str_Parame & "               SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP + HIPCIE_SALCON), HIPCIE_TIPCAM * (HIPCIE_SALCAP + HIPCIE_SALCON))) AS MONTO_TOTAL, "
      g_str_Parame = g_str_Parame & "               SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_PRVGEN + HIPCIE_PRVESP + HIPCIE_PRVCIC), HIPCIE_TIPCAM * (HIPCIE_PRVGEN + HIPCIE_PRVESP + HIPCIE_PRVCIC))) As PROVISIONES "
      g_str_Parame = g_str_Parame & "          FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & "         WHERE HIPCIE_PERANO = " & CStr(moddat_g_int_EdaAno) & " "
      g_str_Parame = g_str_Parame & "           AND HIPCIE_PERMES = " & CStr(r_int_Contad) & " "
      g_str_Parame = g_str_Parame & "           AND HIPCIE_CLAPRV = " & CStr(moddat_g_str_TipPar) & " "
      g_str_Parame = g_str_Parame & "        GROUP BY HIPCIE_CODPRD) "
      g_str_Parame = g_str_Parame & "GROUP BY PRODUCTO "
      g_str_Parame = g_str_Parame & "ORDER BY PRODUCTO "
      
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "Error al ejecutar la consulta de Consolidado de Clasificaciones.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      r_int_NumCol = (r_int_Contad * 3) - 1
      r_int_TotNum = 0
      r_dbl_TotMto = 0
      r_dbl_TotPrv = 0
      
      'Carga grid
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            'Determina producto
            Select Case CInt(g_rst_Princi!PRODUCTO)
               Case 1: r_int_NumFil = 2
               Case 2: r_int_NumFil = 3
               Case 3: r_int_NumFil = 4
               Case 4: r_int_NumFil = 5
            End Select
            
            grd_LisCla.Col = r_int_NumCol
            grd_LisCla.Row = r_int_NumFil
            grd_LisCla.Text = Format(g_rst_Princi!numero, "##,##0")
            If Not IsNull(g_rst_Princi!numero) Then
               r_int_TotNum = r_int_TotNum + g_rst_Princi!numero
            End If
            
            grd_LisCla.Col = r_int_NumCol + 1
            grd_LisCla.Row = r_int_NumFil
            grd_LisCla.Text = Format(g_rst_Princi!MONTO_TOTAL, "###,###,##.00")
            If Not IsNull(g_rst_Princi!MONTO_TOTAL) Then
               r_dbl_TotMto = r_dbl_TotMto + g_rst_Princi!MONTO_TOTAL
            End If
            
            grd_LisCla.Col = r_int_NumCol + 2
            grd_LisCla.Row = r_int_NumFil
            grd_LisCla.Text = Format(g_rst_Princi!PROVISIONES, "###,###,##.00")
            If Not IsNull(g_rst_Princi!PROVISIONES) Then
               r_dbl_TotPrv = r_dbl_TotPrv + g_rst_Princi!PROVISIONES
            End If
            
            r_int_NumFil = r_int_NumFil + 1
            g_rst_Princi.MoveNext
         Loop
         
         'Carga Totales
         grd_LisCla.Col = r_int_NumCol
         grd_LisCla.Row = r_int_NumFil + 1
         grd_LisCla.Text = Format(r_int_TotNum, "##,##0")
         
         grd_LisCla.Col = r_int_NumCol + 1
         grd_LisCla.Row = r_int_NumFil + 1
         grd_LisCla.Text = Format(r_dbl_TotMto, "###,###,##.00")
      
         grd_LisCla.Col = r_int_NumCol + 2
         grd_LisCla.Row = r_int_NumFil + 1
         grd_LisCla.Text = Format(r_dbl_TotPrv, "###,###,##.00")
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next
   
   grd_LisCla.Redraw = True
   Call gs_UbicaGrid(grd_LisCla, 2)
End Sub

Private Sub fs_Buscar_Clasificacion_Det()
Dim r_int_Contad     As Integer
Dim r_int_NumCol     As Integer
Dim r_int_NumFil     As Integer

Dim r_int_CarNum     As Integer
Dim r_dbl_CarMto     As Double
Dim r_dbl_CarPrv     As Double
Dim r_int_AliNum     As Integer
Dim r_dbl_AliMto     As Double
Dim r_dbl_AliPrv     As Double
Dim r_int_TotNum     As Integer
Dim r_dbl_TotMto     As Double
Dim r_dbl_TotPrv     As Double

   grd_LisCla.Redraw = False
   Call gs_LimpiaGrid(grd_LisCla)
   
   'Fila 0
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Row = 0:   grd_LisCla.Text = ""
   grd_LisCla.Col = 1:   grd_LisCla.Text = "PRODUCTOS"
   grd_LisCla.Col = 2:   grd_LisCla.Text = "ENERO"
   grd_LisCla.Col = 3:   grd_LisCla.Text = "ENERO"
   grd_LisCla.Col = 4:   grd_LisCla.Text = "ENERO"
   grd_LisCla.Col = 5:   grd_LisCla.Text = "FEBRERO"
   grd_LisCla.Col = 6:   grd_LisCla.Text = "FEBRERO"
   grd_LisCla.Col = 7:   grd_LisCla.Text = "FEBRERO"
   grd_LisCla.Col = 8:   grd_LisCla.Text = "MARZO"
   grd_LisCla.Col = 9:   grd_LisCla.Text = "MARZO"
   grd_LisCla.Col = 10:  grd_LisCla.Text = "MARZO"
   grd_LisCla.Col = 11:  grd_LisCla.Text = "ABRIL"
   grd_LisCla.Col = 12:  grd_LisCla.Text = "ABRIL"
   grd_LisCla.Col = 13:  grd_LisCla.Text = "ABRIL"
   grd_LisCla.Col = 14:  grd_LisCla.Text = "MAYO"
   grd_LisCla.Col = 15:  grd_LisCla.Text = "MAYO"
   grd_LisCla.Col = 16:  grd_LisCla.Text = "MAYO"
   grd_LisCla.Col = 17:  grd_LisCla.Text = "JUNIO"
   grd_LisCla.Col = 18:  grd_LisCla.Text = "JUNIO"
   grd_LisCla.Col = 19:  grd_LisCla.Text = "JUNIO"
   grd_LisCla.Col = 20:  grd_LisCla.Text = "JULIO"
   grd_LisCla.Col = 21:  grd_LisCla.Text = "JULIO"
   grd_LisCla.Col = 22:  grd_LisCla.Text = "JULIO"
   grd_LisCla.Col = 23:  grd_LisCla.Text = "AGOSTO"
   grd_LisCla.Col = 24:  grd_LisCla.Text = "AGOSTO"
   grd_LisCla.Col = 25:  grd_LisCla.Text = "AGOSTO"
   grd_LisCla.Col = 26:  grd_LisCla.Text = "SETIEMBRE"
   grd_LisCla.Col = 27:  grd_LisCla.Text = "SETIEMBRE"
   grd_LisCla.Col = 28:  grd_LisCla.Text = "SETIEMBRE"
   grd_LisCla.Col = 29:  grd_LisCla.Text = "OCTUBRE"
   grd_LisCla.Col = 30:  grd_LisCla.Text = "OCTUBRE"
   grd_LisCla.Col = 31:  grd_LisCla.Text = "OCTUBRE"
   grd_LisCla.Col = 32:  grd_LisCla.Text = "NOVIEMBRE"
   grd_LisCla.Col = 33:  grd_LisCla.Text = "NOVIEMBRE"
   grd_LisCla.Col = 34:  grd_LisCla.Text = "NOVIEMBRE"
   grd_LisCla.Col = 35:  grd_LisCla.Text = "DICIEMBRE"
   grd_LisCla.Col = 36:  grd_LisCla.Text = "DICIEMBRE"
   grd_LisCla.Col = 37:  grd_LisCla.Text = "DICIEMBRE"
   
   'Fila 1
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 1:   grd_LisCla.Text = "PRODUCTOS"
   grd_LisCla.Col = 2:   grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 3:   grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 4:   grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 5:   grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 6:   grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 7:   grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 8:   grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 9:   grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 10:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 11:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 12:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 13:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 14:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 15:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 16:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 17:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 18:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 19:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 20:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 21:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 22:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 23:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 24:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 25:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 26:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 27:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 28:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 29:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 30:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 31:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 32:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 33:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 34:  grd_LisCla.Text = "PROVISION S/."
   grd_LisCla.Col = 35:  grd_LisCla.Text = "NUMERO"
   grd_LisCla.Col = 36:  grd_LisCla.Text = "MONTO S/."
   grd_LisCla.Col = 37:  grd_LisCla.Text = "PROVISION S/."
   
   'Fila 2
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "CME"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "       CARTERA"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "       ALIENADOS"
   
   'Fila 3
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "CRC-PBP"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "       CARTERA"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "       ALIENADOS"
   
   'Fila 4
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "MICASITA"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "       CARTERA"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "       ALIENADOS"
   
   'Fila 5
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "MIVIVIENDA"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "       CARTERA"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "       ALIENADOS"
   
   'Fila 7
   grd_LisCla.Rows = grd_LisCla.Rows + 2
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "TOTAL CARTERA"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "TOTAL ALIENADOS"
   
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "TOTAL GENERAL"
   
   With grd_LisCla
      .MergeCells = flexMergeFree
      .MergeCol(1) = True
      .MergeRow(0) = True
      .FixedCols = 2
      .FixedRows = 2
   End With
   
   For r_int_Contad = 1 To moddat_g_int_EdaMes
      'Prepara SP que trae consolidado mensual por clasificacion
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT PRODUCTO AS PRODUCTO, SUM(NUMERO) AS NUMERO, SUM(MONTO_TOTAL) AS MONTO_TOTAL, SUM(PROVISIONES) AS PROVISIONES "
      g_str_Parame = g_str_Parame & "  FROM (SELECT CASE WHEN HIPCIE_CODPRD='001' THEN '2' "  'CRC
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='002' THEN '3' "  'MIC
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='003' THEN '1' "  'CME
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='004' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='006' THEN '3' "  'MIC
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='007' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='009' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='010' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='011' THEN '3' "  'MIC
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='012' THEN '3' "  'MIC
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='013' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='014' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='015' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='016' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='017' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='018' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='019' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='021' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='022' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "                    WHEN HIPCIE_CODPRD='023' THEN '4' "  'MIV
      g_str_Parame = g_str_Parame & "               END AS PRODUCTO, "
      g_str_Parame = g_str_Parame & "               COUNT(*) AS NUMERO, "
      g_str_Parame = g_str_Parame & "               SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP + HIPCIE_SALCON), HIPCIE_TIPCAM * (HIPCIE_SALCAP + HIPCIE_SALCON))) AS MONTO_TOTAL, "
      g_str_Parame = g_str_Parame & "               SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_PRVGEN + HIPCIE_PRVESP + HIPCIE_PRVCIC), HIPCIE_TIPCAM * (HIPCIE_PRVGEN + HIPCIE_PRVESP + HIPCIE_PRVCIC))) As PROVISIONES "
      g_str_Parame = g_str_Parame & "          FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & "         WHERE HIPCIE_PERANO = " & CStr(moddat_g_int_EdaAno) & " "
      g_str_Parame = g_str_Parame & "           AND HIPCIE_PERMES = " & CStr(r_int_Contad) & " "
      g_str_Parame = g_str_Parame & "           AND HIPCIE_CLAPRV = " & CStr(moddat_g_str_TipPar) & " "
      g_str_Parame = g_str_Parame & "        GROUP BY HIPCIE_CODPRD) "
      g_str_Parame = g_str_Parame & "GROUP BY PRODUCTO "
      g_str_Parame = g_str_Parame & "ORDER BY PRODUCTO "
      
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "Error al ejecutar la consulta de Consolidado de Clasificaciones.", vbCritical, modgen_g_str_NomPlt
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      r_int_NumCol = (r_int_Contad * 3) - 1
      r_int_CarNum = 0
      r_dbl_CarMto = 0
      r_dbl_CarPrv = 0
      r_int_AliNum = 0
      r_dbl_AliMto = 0
      r_dbl_AliPrv = 0
      r_int_TotNum = 0
      r_dbl_TotMto = 0
      r_dbl_TotPrv = 0
      
      'Carga grid
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            'Determina producto
            Select Case CInt(g_rst_Princi!PRODUCTO)
               Case 1: r_int_NumFil = 2
               Case 2: r_int_NumFil = 5
               Case 3: r_int_NumFil = 8
               Case 4: r_int_NumFil = 11
            End Select
            
            grd_LisCla.Col = r_int_NumCol
            grd_LisCla.Row = r_int_NumFil
            grd_LisCla.Text = Format(g_rst_Princi!numero, "##,##0")
            If Not IsNull(g_rst_Princi!numero) Then
               r_int_TotNum = r_int_TotNum + g_rst_Princi!numero
            End If
            
            grd_LisCla.Col = r_int_NumCol + 1
            grd_LisCla.Row = r_int_NumFil
            grd_LisCla.Text = Format(g_rst_Princi!MONTO_TOTAL, "###,###,##.00")
            If Not IsNull(g_rst_Princi!MONTO_TOTAL) Then
               r_dbl_TotMto = r_dbl_TotMto + g_rst_Princi!MONTO_TOTAL
            End If
            
            grd_LisCla.Col = r_int_NumCol + 2
            grd_LisCla.Row = r_int_NumFil
            grd_LisCla.Text = Format(g_rst_Princi!PROVISIONES, "###,###,##.00")
            If Not IsNull(g_rst_Princi!PROVISIONES) Then
               r_dbl_TotPrv = r_dbl_TotPrv + g_rst_Princi!PROVISIONES
            End If
               
            g_rst_Princi.MoveNext
         Loop
         
         r_int_NumFil = r_int_NumFil + 3
         
         'Carga Totales
         grd_LisCla.Col = r_int_NumCol
         grd_LisCla.Row = r_int_NumFil + 1
         grd_LisCla.Text = Format(r_int_TotNum, "##,##0")
         
         grd_LisCla.Col = r_int_NumCol + 1
         grd_LisCla.Row = r_int_NumFil + 1
         grd_LisCla.Text = Format(r_dbl_TotMto, "###,###,##.00")
      
         grd_LisCla.Col = r_int_NumCol + 2
         grd_LisCla.Row = r_int_NumFil + 1
         grd_LisCla.Text = Format(r_dbl_TotPrv, "###,###,##.00")
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next
   
   grd_LisCla.Redraw = True
   Call gs_UbicaGrid(grd_LisCla, 2)
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_nrofil        As Integer
Dim r_int_NoFlLi        As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      'Titulo
      .Cells(1, 1) = "CONSOLIDADO DE CLASIFICACION " & UCase(moddat_g_str_NomPrd) & " EN SOLES"
      .Range(.Cells(1, 1), .Cells(1, 37)).Merge
      .Range("A1:Y1").HorizontalAlignment = xlHAlignCenter
      
      'Primera Linea
      r_int_nrofil = 3
      .Cells(r_int_nrofil, 1) = "PRODUCTOS"
      .Cells(r_int_nrofil, 2) = "ENERO"
      .Cells(r_int_nrofil, 5) = "FEBRERO"
      .Cells(r_int_nrofil, 8) = "MARZO"
      .Cells(r_int_nrofil, 11) = "ABRIL"
      .Cells(r_int_nrofil, 14) = "MAYO"
      .Cells(r_int_nrofil, 17) = "JUNIO"
      .Cells(r_int_nrofil, 20) = "JULIO"
      .Cells(r_int_nrofil, 23) = "AGOSTO"
      .Cells(r_int_nrofil, 26) = "SETIEMBRE"
      .Cells(r_int_nrofil, 29) = "OCTUBRE"
      .Cells(r_int_nrofil, 32) = "NOVIEMBRE"
      .Cells(r_int_nrofil, 35) = "DICIEMBRE"
      .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 4)).Merge
      .Range(.Cells(r_int_nrofil, 5), .Cells(r_int_nrofil, 7)).Merge
      .Range(.Cells(r_int_nrofil, 8), .Cells(r_int_nrofil, 10)).Merge
      .Range(.Cells(r_int_nrofil, 11), .Cells(r_int_nrofil, 13)).Merge
      .Range(.Cells(r_int_nrofil, 14), .Cells(r_int_nrofil, 16)).Merge
      .Range(.Cells(r_int_nrofil, 17), .Cells(r_int_nrofil, 19)).Merge
      .Range(.Cells(r_int_nrofil, 20), .Cells(r_int_nrofil, 22)).Merge
      .Range(.Cells(r_int_nrofil, 23), .Cells(r_int_nrofil, 25)).Merge
      .Range(.Cells(r_int_nrofil, 26), .Cells(r_int_nrofil, 28)).Merge
      .Range(.Cells(r_int_nrofil, 29), .Cells(r_int_nrofil, 31)).Merge
      .Range(.Cells(r_int_nrofil, 32), .Cells(r_int_nrofil, 34)).Merge
      .Range(.Cells(r_int_nrofil, 35), .Cells(r_int_nrofil, 37)).Merge
      
      'Segunda Linea
      r_int_nrofil = r_int_nrofil + 1
      .Columns("A").ColumnWidth = 15
      .Columns("B").ColumnWidth = 9:    .Cells(r_int_nrofil, 2) = "NUMERO":          .Cells(r_int_nrofil, 2).HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 13:   .Cells(r_int_nrofil, 3) = "MONTO S/.":       .Cells(r_int_nrofil, 3).HorizontalAlignment = xlHAlignRight
      .Columns("D").ColumnWidth = 13:   .Cells(r_int_nrofil, 4) = "PROVISION S/.":   .Cells(r_int_nrofil, 4).HorizontalAlignment = xlHAlignRight
      .Columns("E").ColumnWidth = 9:    .Cells(r_int_nrofil, 5) = "NUMERO":          .Cells(r_int_nrofil, 5).HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 13:   .Cells(r_int_nrofil, 6) = "MONTO":           .Cells(r_int_nrofil, 6).HorizontalAlignment = xlHAlignRight
      .Columns("G").ColumnWidth = 13:   .Cells(r_int_nrofil, 7) = "PROVISION S/.":   .Cells(r_int_nrofil, 7).HorizontalAlignment = xlHAlignRight
      .Columns("H").ColumnWidth = 9:    .Cells(r_int_nrofil, 8) = "NUMERO":          .Cells(r_int_nrofil, 8).HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 13:   .Cells(r_int_nrofil, 9) = "MONTO":           .Cells(r_int_nrofil, 9).HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 13:   .Cells(r_int_nrofil, 10) = "PROVISION S/.":  .Cells(r_int_nrofil, 10).HorizontalAlignment = xlHAlignRight
      .Columns("K").ColumnWidth = 9:    .Cells(r_int_nrofil, 11) = "NUMERO":         .Cells(r_int_nrofil, 11).HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 13:   .Cells(r_int_nrofil, 12) = "MONTO":          .Cells(r_int_nrofil, 12).HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 13:   .Cells(r_int_nrofil, 13) = "PROVISION S/.":  .Cells(r_int_nrofil, 13).HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 9:    .Cells(r_int_nrofil, 14) = "NUMERO":         .Cells(r_int_nrofil, 14).HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 13:   .Cells(r_int_nrofil, 15) = "MONTO":          .Cells(r_int_nrofil, 15).HorizontalAlignment = xlHAlignRight
      .Columns("P").ColumnWidth = 13:   .Cells(r_int_nrofil, 16) = "PROVISION S/.":  .Cells(r_int_nrofil, 16).HorizontalAlignment = xlHAlignRight
      .Columns("Q").ColumnWidth = 9:    .Cells(r_int_nrofil, 17) = "NUMERO":         .Cells(r_int_nrofil, 17).HorizontalAlignment = xlHAlignCenter
      .Columns("R").ColumnWidth = 13:   .Cells(r_int_nrofil, 18) = "MONTO":          .Cells(r_int_nrofil, 18).HorizontalAlignment = xlHAlignRight
      .Columns("S").ColumnWidth = 13:   .Cells(r_int_nrofil, 19) = "PROVISION S/.":  .Cells(r_int_nrofil, 19).HorizontalAlignment = xlHAlignRight
      .Columns("T").ColumnWidth = 9:    .Cells(r_int_nrofil, 20) = "NUMERO":         .Cells(r_int_nrofil, 20).HorizontalAlignment = xlHAlignCenter
      .Columns("U").ColumnWidth = 13:   .Cells(r_int_nrofil, 21) = "MONTO":          .Cells(r_int_nrofil, 21).HorizontalAlignment = xlHAlignRight
      .Columns("V").ColumnWidth = 13:   .Cells(r_int_nrofil, 22) = "PROVISION S/.":  .Cells(r_int_nrofil, 22).HorizontalAlignment = xlHAlignRight
      .Columns("W").ColumnWidth = 9:    .Cells(r_int_nrofil, 23) = "NUMERO":         .Cells(r_int_nrofil, 23).HorizontalAlignment = xlHAlignCenter
      .Columns("X").ColumnWidth = 13:   .Cells(r_int_nrofil, 24) = "MONTO":          .Cells(r_int_nrofil, 24).HorizontalAlignment = xlHAlignRight
      .Columns("Y").ColumnWidth = 13:   .Cells(r_int_nrofil, 25) = "PROVISION S/.":  .Cells(r_int_nrofil, 25).HorizontalAlignment = xlHAlignRight
      .Columns("Z").ColumnWidth = 9:    .Cells(r_int_nrofil, 26) = "NUMERO":         .Cells(r_int_nrofil, 26).HorizontalAlignment = xlHAlignCenter
      .Columns("AA").ColumnWidth = 13:  .Cells(r_int_nrofil, 27) = "MONTO":          .Cells(r_int_nrofil, 27).HorizontalAlignment = xlHAlignRight
      .Columns("AB").ColumnWidth = 13:  .Cells(r_int_nrofil, 28) = "PROVISION S/.":  .Cells(r_int_nrofil, 28).HorizontalAlignment = xlHAlignRight
      .Columns("AC").ColumnWidth = 9:   .Cells(r_int_nrofil, 29) = "NUMERO":         .Cells(r_int_nrofil, 29).HorizontalAlignment = xlHAlignCenter
      .Columns("AD").ColumnWidth = 13:  .Cells(r_int_nrofil, 30) = "MONTO":          .Cells(r_int_nrofil, 30).HorizontalAlignment = xlHAlignRight
      .Columns("AE").ColumnWidth = 13:  .Cells(r_int_nrofil, 31) = "PROVISION S/.":  .Cells(r_int_nrofil, 31).HorizontalAlignment = xlHAlignRight
      .Columns("AF").ColumnWidth = 9:   .Cells(r_int_nrofil, 32) = "NUMERO":         .Cells(r_int_nrofil, 32).HorizontalAlignment = xlHAlignCenter
      .Columns("AG").ColumnWidth = 13:  .Cells(r_int_nrofil, 33) = "MONTO":          .Cells(r_int_nrofil, 33).HorizontalAlignment = xlHAlignRight
      .Columns("AH").ColumnWidth = 13:  .Cells(r_int_nrofil, 34) = "PROVISION S/.":  .Cells(r_int_nrofil, 34).HorizontalAlignment = xlHAlignRight
      .Columns("AI").ColumnWidth = 9:   .Cells(r_int_nrofil, 35) = "NUMERO":         .Cells(r_int_nrofil, 35).HorizontalAlignment = xlHAlignCenter
      .Columns("AJ").ColumnWidth = 13:  .Cells(r_int_nrofil, 36) = "MONTO":          .Cells(r_int_nrofil, 36).HorizontalAlignment = xlHAlignRight
      .Columns("AK").ColumnWidth = 13:  .Cells(r_int_nrofil, 37) = "PROVISION S/.":  .Cells(r_int_nrofil, 37).HorizontalAlignment = xlHAlignRight
      
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
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 37)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 37)).Font.Size = 11
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 37)).Font.Bold = True
      
      'Exporta filas
      For r_int_Contad = 5 To 10
         .Cells(r_int_Contad, 1) = grd_LisCla.TextMatrix(r_int_Contad - 3, 1)
         .Cells(r_int_Contad, 2) = grd_LisCla.TextMatrix(r_int_Contad - 3, 2)
         .Cells(r_int_Contad, 3) = grd_LisCla.TextMatrix(r_int_Contad - 3, 3)
         .Cells(r_int_Contad, 4) = grd_LisCla.TextMatrix(r_int_Contad - 3, 4)
         .Cells(r_int_Contad, 5) = grd_LisCla.TextMatrix(r_int_Contad - 3, 5)
         .Cells(r_int_Contad, 6) = grd_LisCla.TextMatrix(r_int_Contad - 3, 6)
         .Cells(r_int_Contad, 7) = grd_LisCla.TextMatrix(r_int_Contad - 3, 7)
         .Cells(r_int_Contad, 8) = grd_LisCla.TextMatrix(r_int_Contad - 3, 8)
         .Cells(r_int_Contad, 9) = grd_LisCla.TextMatrix(r_int_Contad - 3, 9)
         .Cells(r_int_Contad, 10) = grd_LisCla.TextMatrix(r_int_Contad - 3, 10)
         .Cells(r_int_Contad, 11) = grd_LisCla.TextMatrix(r_int_Contad - 3, 11)
         .Cells(r_int_Contad, 12) = grd_LisCla.TextMatrix(r_int_Contad - 3, 12)
         .Cells(r_int_Contad, 13) = grd_LisCla.TextMatrix(r_int_Contad - 3, 13)
         .Cells(r_int_Contad, 14) = grd_LisCla.TextMatrix(r_int_Contad - 3, 14)
         .Cells(r_int_Contad, 15) = grd_LisCla.TextMatrix(r_int_Contad - 3, 15)
         .Cells(r_int_Contad, 16) = grd_LisCla.TextMatrix(r_int_Contad - 3, 16)
         .Cells(r_int_Contad, 17) = grd_LisCla.TextMatrix(r_int_Contad - 3, 17)
         .Cells(r_int_Contad, 18) = grd_LisCla.TextMatrix(r_int_Contad - 3, 18)
         .Cells(r_int_Contad, 19) = grd_LisCla.TextMatrix(r_int_Contad - 3, 19)
         .Cells(r_int_Contad, 20) = grd_LisCla.TextMatrix(r_int_Contad - 3, 20)
         .Cells(r_int_Contad, 21) = grd_LisCla.TextMatrix(r_int_Contad - 3, 21)
         .Cells(r_int_Contad, 22) = grd_LisCla.TextMatrix(r_int_Contad - 3, 22)
         .Cells(r_int_Contad, 23) = grd_LisCla.TextMatrix(r_int_Contad - 3, 23)
         .Cells(r_int_Contad, 24) = grd_LisCla.TextMatrix(r_int_Contad - 3, 24)
         .Cells(r_int_Contad, 25) = grd_LisCla.TextMatrix(r_int_Contad - 3, 25)
         .Cells(r_int_Contad, 26) = grd_LisCla.TextMatrix(r_int_Contad - 3, 26)
         .Cells(r_int_Contad, 27) = grd_LisCla.TextMatrix(r_int_Contad - 3, 27)
         .Cells(r_int_Contad, 28) = grd_LisCla.TextMatrix(r_int_Contad - 3, 28)
         .Cells(r_int_Contad, 29) = grd_LisCla.TextMatrix(r_int_Contad - 3, 29)
         .Cells(r_int_Contad, 30) = grd_LisCla.TextMatrix(r_int_Contad - 3, 30)
         .Cells(r_int_Contad, 31) = grd_LisCla.TextMatrix(r_int_Contad - 3, 31)
         .Cells(r_int_Contad, 32) = grd_LisCla.TextMatrix(r_int_Contad - 3, 32)
         .Cells(r_int_Contad, 33) = grd_LisCla.TextMatrix(r_int_Contad - 3, 33)
         .Cells(r_int_Contad, 34) = grd_LisCla.TextMatrix(r_int_Contad - 3, 34)
         .Cells(r_int_Contad, 35) = grd_LisCla.TextMatrix(r_int_Contad - 3, 35)
         .Cells(r_int_Contad, 36) = grd_LisCla.TextMatrix(r_int_Contad - 3, 36)
         .Cells(r_int_Contad, 37) = grd_LisCla.TextMatrix(r_int_Contad - 3, 37)
      Next
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

