VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Pro_CtbPpg_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10275
   Icon            =   "GesCtb_frm_913.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel2 
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   10320
      _Version        =   65536
      _ExtentX        =   18203
      _ExtentY        =   1508
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
      Begin Threed.SSPanel pnl_NroOperacion 
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   120
         Width           =   2265
         _Version        =   65536
         _ExtentX        =   3995
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   4210752
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
         Left            =   1320
         TabIndex        =   15
         Top             =   450
         Width           =   5985
         _Version        =   65536
         _ExtentX        =   10557
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   4210752
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
      Begin Threed.SSPanel pnl_DNI 
         Height          =   315
         Left            =   8040
         TabIndex        =   16
         Top             =   450
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   556
         _StockProps     =   15
         ForeColor       =   4210752
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nro Operación:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   165
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "DNI:"
         Height          =   195
         Left            =   7560
         TabIndex        =   11
         Top             =   525
         Width           =   330
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   525
         Width           =   600
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      _Version        =   65536
      _ExtentX        =   18203
      _ExtentY        =   1085
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
         Height          =   585
         Left            =   630
         TabIndex        =   1
         Top             =   30
         Width           =   3765
         _Version        =   65536
         _ExtentX        =   6641
         _ExtentY        =   1032
         _StockProps     =   15
         Caption         =   "Consulta del detalle de la operación"
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
         Index           =   1
         Left            =   50
         Picture         =   "GesCtb_frm_913.frx":000C
         Top             =   120
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   3750
      Left            =   0
      TabIndex        =   2
      Top             =   2200
      Width           =   10320
      _Version        =   65536
      _ExtentX        =   18203
      _ExtentY        =   6615
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
      Begin Threed.SSPanel pnl_Haber 
         Height          =   315
         Left            =   8670
         TabIndex        =   20
         Top             =   3385
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2258
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00000 "
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
      Begin Threed.SSPanel pnl_Debe 
         Height          =   315
         Left            =   7380
         TabIndex        =   13
         Top             =   3385
         Width           =   1280
         _Version        =   65536
         _ExtentX        =   2258
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00000 "
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
      Begin Threed.SSPanel pnl_Tit_NumOpe 
         Height          =   285
         Left            =   70
         TabIndex        =   4
         Top             =   90
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "N° Operación"
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
      Begin Threed.SSPanel pnl_Tit_TipPpg 
         Height          =   285
         Left            =   5280
         TabIndex        =   5
         Top             =   90
         Width           =   2150
         _Version        =   65536
         _ExtentX        =   3792
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Nro Cuenta"
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
      Begin Threed.SSPanel pnl_Tit_FecPro 
         Height          =   285
         Left            =   8670
         TabIndex        =   6
         Top             =   90
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2205
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Haber"
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
         Left            =   1300
         TabIndex        =   7
         Top             =   90
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
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
      Begin Threed.SSPanel pnl_Tit_FecPpg 
         Height          =   285
         Left            =   7420
         TabIndex        =   8
         Top             =   90
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2205
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "Debe"
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
         Height          =   3030
         Left            =   45
         TabIndex        =   3
         Top             =   405
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   5345
         _Version        =   393216
         Rows            =   11
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         BackColorSel    =   32768
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin Threed.SSPanel SSPanel10 
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   660
      Width           =   10320
      _Version        =   65536
      _ExtentX        =   18203
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
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   9600
         Picture         =   "GesCtb_frm_913.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   45
         Picture         =   "GesCtb_frm_913.frx":0758
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Exportar a Excel"
         Top             =   30
         Width           =   585
      End
   End
End
Attribute VB_Name = "frm_Pro_CtbPpg_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
     
   Call fs_IniciaGrid
   Call fs_Limpiar
   Call fs_Buscar_operacion
   
   Call gs_CentraForm(Me)
   Call gs_RefrescaGrid(grd_Listad)
   Screen.MousePointer = 0
End Sub

Private Sub fs_IniciaGrid()
   grd_Listad.ColWidth(0) = 1230 'codigo del producto
   grd_Listad.ColWidth(1) = 3975 'Glosa
   grd_Listad.ColWidth(2) = 2150 'Nro de Cuenta
   grd_Listad.ColWidth(3) = 1245 'debe
   grd_Listad.ColWidth(4) = 1245 'haber
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
   
   pnl_Debe.Caption = "0.00000"
   pnl_Haber.Caption = "0.00000"
   grd_Listad.Rows = 0
End Sub

Private Sub fs_Limpiar()
    pnl_NomCli.Caption = ""
    pnl_NroOperacion.Caption = ""
    pnl_DNI.Caption = ""
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "CONTABILIZACION DE ASIENTOS DEL PREPAGO"
      .Range(.Cells(2, 2), .Cells(2, 7)).Merge
      .Range(.Cells(2, 2), .Cells(2, 7)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 7)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(5, 2) = "NRO. DE OPERACION"
      .Cells(5, 3) = "GLOSA"
      .Cells(5, 4) = "Nro CUENTA"
      .Cells(5, 5) = "DEBE"
      .Cells(5, 6) = "HABER"
      
      .Range(.Cells(5, 2), .Cells(5, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(5, 2), .Cells(5, 6)).Font.Bold = True
      .Range(.Cells(5, 3), .Cells(5, 6)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 20
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 35
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 20
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignRight
      .Columns("E").NumberFormat = "###,###,##0.00"
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignRight
      .Columns("F").NumberFormat = "###,###,##0.00"
      
      .Range(.Cells(1, 1), .Cells(10, 6)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 6)).Font.Size = 11
      
      Dim swtDebe As Double
      Dim swtHaber As Double
      
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil + 3, 2) = "'" & CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 0)) 'N°Operacion
         .Cells(r_int_NumFil + 3, 3) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 1)) 'Glosa
         .Cells(r_int_NumFil + 3, 4) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 3, 2) 'NroCuenta
         .Cells(r_int_NumFil + 3, 5) = grd_Listad.TextMatrix(r_int_NumFil - 3, 3) 'Debe
         .Cells(r_int_NumFil + 3, 6) = grd_Listad.TextMatrix(r_int_NumFil - 3, 4) 'Haber
         
         swtDebe = swtDebe + grd_Listad.TextMatrix(r_int_NumFil - 3, 3)
         swtHaber = swtHaber + grd_Listad.TextMatrix(r_int_NumFil - 3, 4)
         
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
            
      .Cells(r_int_NumFil + 3, 5).Interior.Color = RGB(146, 208, 80)
      .Cells(r_int_NumFil + 3, 6).Interior.Color = RGB(146, 208, 80)
      .Cells(r_int_NumFil + 3, 5).Font.Bold = True
      .Cells(r_int_NumFil + 3, 6).Font.Bold = True
      .Cells(r_int_NumFil + 3, 5) = swtDebe
      .Cells(r_int_NumFil + 3, 6) = swtHaber
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Buscar_operacion()
Dim swtNumOper       As String
Dim r_int_TipPgo     As Integer
Dim r_str_Glosa      As String

   pnl_NroOperacion.Caption = Trim(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = Trim(moddat_g_str_NomCli)
   pnl_DNI.Caption = Trim(moddat_g_str_NumDoc)
   swtNumOper = Left(moddat_g_str_NumOpe, 3) & Mid(moddat_g_str_NumOpe, 5, 2) & Right(moddat_g_str_NumOpe, 5)

   Set g_rst_Princi = Nothing

   'DATOS DE LA OPERACION DEL PREPAGO
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PPGCAB_MTODEP, PPGCAB_MTOAPL, PPGCAB_SEGDES, "
   g_str_Parame = g_str_Parame & " PPGCAB_SEGINM, PPGCAB_INTCAL_TNC, PPGCAB_INTCAL_TC, "
   g_str_Parame = g_str_Parame & " PPGCAB_PBPPER, PPGCAB_PBPINT, PPGCAB_MTOPOR,"
   g_str_Parame = g_str_Parame & " PPGCAB_MTOITF, PPGCAB_TIPPPG, CH.hipmae_moneda,"
   g_str_Parame = g_str_Parame & " PPGCAB_MTOTOT, PPGCAB_SLDACT_TNC, PPGCAB_SLDACT_TC "
   g_str_Parame = g_str_Parame & " FROM CRE_PPGCAB PP "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE CH ON CH.HIPMAE_NUMOPE = PP.PPGCAB_NUMOPE "
   g_str_Parame = g_str_Parame & " WHERE PPGCAB_NUMOPE = '" & swtNumOper & "'  "
   g_str_Parame = g_str_Parame & " AND PPGCAB_FECPPG =  " & moddat_g_str_FecIng & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   'Tipo de prepago (parcial = 1 o total = 2)
   r_int_TipPgo = Trim(g_rst_Princi!PPGCAB_TIPPPG)
   If r_int_TipPgo = 1 Then
      r_str_Glosa = "PPG PAR - " & Trim(swtNumOper) & " - "
   Else
      r_str_Glosa = "PPG TOT - " & Trim(swtNumOper) & " - "
   End If
   
   Dim r_str_Portes As String, r_str_IntTNC As String, r_str_IntTC As String, r_str_SegDesg As String
   Dim r_str_SegInm As String, r_str_CapPBP As String, r_str_IntPBP As String, r_str_ITF As String
   Dim r_str_MtoApl As String, r_str_MtoDep As String, r_str_Codigo As String '
   
   r_str_Codigo = Left(swtNumOper, 3)
   r_str_Portes = "521229010109"
   r_str_SegDesg = "251602010103"
   r_str_ITF = "251705010107"
   
   Select Case r_str_Codigo
      'CREDITO CRC-PBP
      Case "001"
         r_str_IntTNC = "512401042401"
         r_str_IntTC = "512401042401"
         r_str_SegInm = "252602010104"
         r_str_CapPBP = "142104240101"
         r_str_IntPBP = "512401042401"
         r_str_MtoApl = "142104240101"
         r_str_MtoDep = "112301060202"
      
      'CREDITO MICASITA DOLARES
      Case "002"
         r_str_IntTNC = "512401040601"
         r_str_IntTC = "512401040601"
         r_str_SegInm = "252602010104"
         r_str_CapPBP = "142104060101"
         r_str_IntPBP = "512401040601"
         r_str_MtoApl = "142104060101"
         r_str_MtoDep = "112301060202"
         r_str_SegDesg = "252602010103"
      
      'CREDITO CME
      Case "003"
         r_str_IntTNC = "511401042501"
         r_str_IntTC = "511401042501"
         r_str_SegInm = "251602010103"
         r_str_CapPBP = "141104250101"
         r_str_IntPBP = "511401042501"
         r_str_MotApl = "141104250101"
         r_str_MtoDep = "111301060201"
      
      'CREDITO PROYECTO MIHOGAR
      Case "004"
         r_str_IntTNC = "511401042301"
         r_str_IntTC = "511401042301"
         r_str_SegInm = "251602010104"
         r_str_CapPBP = "141104230101"
         r_str_IntPBP = "511401042301"
         r_str_MtoApl = "141104230101"
         r_str_MtoDep = "111301060201"
      
      'CREDITO FMV UNION ANDINA
      Case "009"
         r_str_IntTNC = "511401042305"
         r_str_IntTC = "511401042305"
         r_str_SegInm = "251602010104"
         r_str_CapPBP = "141104230103"
         r_str_IntPBP = "511401042305"
         r_str_MtoApl = "141104230103"
         r_str_MtoDep = "111301060201"
      
      'CREDITO MICASITA SOLES
      Case "011", "006"
         r_str_IntTNC = "511401042302"
         r_str_IntTC = "511401042302"
         r_str_SegInm = "251602010104"
         r_str_CapPBP = "141104060101"
         r_str_IntPBP = "511401042302"
         r_str_MtoApl = "141104060101"
         r_str_MtoDep = "111301060201"
      
      'CREDITO MIVIVIENDA
      Case "007", "010", "013", "014", "015", "016", "017", "018", "019", "021", "022", "023"
         r_str_IntTNC = "511401042302"
         r_str_IntTC = "511401042302"
         r_str_SegInm = "251602010104"
         r_str_CapPBP = "141104230102"
         r_str_IntPBP = "511401042302"
         r_str_MtoApl = "141104230102"
         r_str_MtoDep = "111301060201"
   End Select
   '-----------------------------------------------------------------------------
   
   If ((g_rst_Princi!PPGCAB_MTODEP <> Null) Or (g_rst_Princi!PPGCAB_MTODEP > 0)) Then
      'Monto Depositado - Debe
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "MTO DEP"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_MtoDep
       grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_MTODEP, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(0, "###,###,##0.00")
   End If
   If ((r_int_TipPgo = 2) And ((g_rst_Princi!PPGCAB_MTOTOT <> Null) Or (g_rst_Princi!PPGCAB_MTOTOT > 0))) Then
      'Monto del Prepago Total - Debe
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "MTO DEP"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_MtoDep 'r_str_MtoApl
       grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_MTOTOT, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(0, "###,###,##0.00")
   End If
   If (g_rst_Princi!PPGCAB_MTOPOR <> Null) Or (g_rst_Princi!PPGCAB_MTOPOR > 0) Then
      'Portes - Haber
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "PORTES"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_Portes
       grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_MTOPOR, "###,###,##0.00")
   End If
   If (g_rst_Princi!PPGCAB_INTCAL_TNC <> Null) Or (g_rst_Princi!PPGCAB_INTCAL_TNC > 0) Then
      'Intereses TNC a la fecha - Haber
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "INT TNC"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_IntTNC
       grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_INTCAL_TNC, "###,###,##0.00")
   End If
   If (g_rst_Princi!PPGCAB_INTCAL_TC <> Null) Or (g_rst_Princi!PPGCAB_INTCAL_TC > 0) Then
      'Intereses TC a la fecha - Haber
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "INT TC"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_IntTC
       grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_INTCAL_TC, "###,###,##0.00")
   End If
   If (g_rst_Princi!PPGCAB_SEGDES <> Null) Or (g_rst_Princi!PPGCAB_SEGDES > 0) Then
      'Seguro Desgravamen - Haber
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "SEG DES"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_SegDesg
       grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_SEGDES, "###,###,##0.00")
   End If
   If (g_rst_Princi!PPGCAB_SEGINM <> Null) Or (g_rst_Princi!PPGCAB_SEGINM > 0) Then
      'Seguro Inmueble - Haber
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "SEG INM"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_SegInm
       grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_SEGINM, "###,###,##0.00")
   End If
   If (g_rst_Princi!PPGCAB_PBPPER <> Null) Or (g_rst_Princi!PPGCAB_PBPPER > 0) Then
      'Capital PBP - Haber
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "CAP PBP"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_CapPBP
       grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_PBPPER, "###,###,##0.00")
   End If
   If (g_rst_Princi!PPGCAB_PBPINT <> Null) Or (g_rst_Princi!PPGCAB_PBPINT > 0) Then
      'Interes PBP - Haber
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "INT PBP"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_IntPBP
       grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_PBPINT, "###,###,##0.00")
   End If
   If (g_rst_Princi!PPGCAB_MTOITF <> Null) Or (g_rst_Princi!PPGCAB_MTOITF > 0) Then
      'ITF - Haber
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "ITF"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_ITF
       grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_MTOITF, "###,###,##0.00")
   End If
   If ((r_int_TipPgo = 1) And ((g_rst_Princi!PPGCAB_MTOAPL <> Null) Or (g_rst_Princi!PPGCAB_MTOAPL > 0))) Then
      'Monto de Prepago a Aplicar - Haber
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
       grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "MTO APL"
       grd_Listad.Col = 2: grd_Listad.Text = r_str_MtoApl
       grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
       grd_Listad.Col = 4: grd_Listad.Text = Format(g_rst_Princi!PPGCAB_MTOAPL, "###,###,##0.00")
   End If
   '---------------------------------------------------------
   If (r_int_TipPgo = 2) Then
       Dim swtSaldo As Double
       swtSaldo = CDbl(g_rst_Princi!PPGCAB_SLDACT_TNC) + CDbl(g_rst_Princi!PPGCAB_SLDACT_TC)
       If (swtSaldo > 0) Then
          'Saldo Actual - Haber
           grd_Listad.Rows = grd_Listad.Rows + 1
           grd_Listad.Row = grd_Listad.Rows - 1
           grd_Listad.Col = 0: grd_Listad.Text = Trim(moddat_g_str_NumOpe)
           grd_Listad.Col = 1: grd_Listad.Text = r_str_Glosa & "MTO APL" '"SALDO"
           grd_Listad.Col = 2: grd_Listad.Text = r_str_MtoApl 'r_str_Saldo
           grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
           grd_Listad.Col = 4: grd_Listad.Text = Format(swtSaldo, "###,###,##0.00")
       End If
   End If
   
   Call Sumar_Columnas
End Sub

Private Sub Sumar_Columnas()
Dim swtTotDebe As Double
Dim swtTotHaber As Double
    
    For r_Fila = 0 To grd_Listad.Rows - 1
        swtTotDebe = swtTotDebe + grd_Listad.TextMatrix(r_Fila, 3)
        swtTotHaber = swtTotHaber + grd_Listad.TextMatrix(r_Fila, 4)
    Next r_Fila
    
    pnl_Debe.Caption = Format(swtTotDebe, "###,###,##0.00")
    pnl_Haber.Caption = Format(swtTotHaber, "###,###,##0.00")
End Sub
