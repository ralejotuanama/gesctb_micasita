VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_CtbPbp_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9030
   Icon            =   "GesCtb_frm_915.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9030
      _Version        =   65536
      _ExtentX        =   15928
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   225
         Left            =   630
         TabIndex        =   1
         Top             =   200
         Width           =   3165
         _Version        =   65536
         _ExtentX        =   5583
         _ExtentY        =   397
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
         Left            =   45
         Picture         =   "GesCtb_frm_915.frx":000C
         Top             =   90
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   3870
      Left            =   0
      TabIndex        =   2
      Top             =   2160
      Width           =   9030
      _Version        =   65536
      _ExtentX        =   15928
      _ExtentY        =   6826
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
         Left            =   7485
         TabIndex        =   3
         Top             =   3460
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
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
         Left            =   6195
         TabIndex        =   4
         Top             =   3465
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
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
      Begin Threed.SSPanel pnl_Tit_TipPpg 
         Height          =   285
         Left            =   4050
         TabIndex        =   5
         Top             =   120
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
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
         Left            =   7485
         TabIndex        =   6
         Top             =   120
         Width           =   1290
         _Version        =   65536
         _ExtentX        =   2284
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
         Left            =   80
         TabIndex        =   7
         Top             =   120
         Width           =   3990
         _Version        =   65536
         _ExtentX        =   7038
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
         Left            =   6195
         TabIndex        =   8
         Top             =   120
         Width           =   1300
         _Version        =   65536
         _ExtentX        =   2293
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
         TabIndex        =   9
         Top             =   405
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   5345
         _Version        =   393216
         Rows            =   11
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         BackColorSel    =   32768
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   630
      Width           =   9030
      _Version        =   65536
      _ExtentX        =   15928
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
      Begin VB.CommandButton cmd_Detalle 
         Height          =   585
         Left            =   645
         Picture         =   "GesCtb_frm_915.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ver Detalle"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   45
         Picture         =   "GesCtb_frm_915.frx":0758
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exportar a Excel"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   8370
         Picture         =   "GesCtb_frm_915.frx":0A62
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   1290
      Width           =   9030
      _Version        =   65536
      _ExtentX        =   15928
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
      Begin Threed.SSPanel pnl_periodo 
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
      Begin Threed.SSPanel pnl_NomProd 
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   525
         Width           =   690
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Periodo:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   165
         Width           =   585
      End
   End
End
Attribute VB_Name = "frm_Pro_CtbPbp_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Detalle_Click()
   If grd_Listad.Rows = -1 Then
      Exit Sub
   End If
   
   If (CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 2)) + CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 3))) = 0 Then
      Exit Sub
   End If
   
   moddat_g_str_NomPrd = Trim(pnl_NomProd.Caption)
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_Pro_CtbPbp_03.Show 1
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
     
   Call fs_IniciaGrid
   Call fs_Limpiar
   Call fs_Buscar_detalle
   
   Call gs_CentraForm(Me)
   Call gs_RefrescaGrid(grd_Listad)
   Screen.MousePointer = 0
End Sub

Private Sub fs_IniciaGrid()
   grd_Listad.ColWidth(0) = 3975 'Glosa
   grd_Listad.ColWidth(1) = 2150 'Nro de Cuenta
   grd_Listad.ColWidth(2) = 1245 'debe
   grd_Listad.ColWidth(3) = 1245 'haber
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   
   pnl_Debe.Caption = "0.00000" & " "
   pnl_Haber.Caption = "0.00000" & " "
   grd_Listad.Rows = 0
End Sub

Private Sub fs_Limpiar()
   pnl_NomProd.Caption = ""
   pnl_Periodo.Caption = ""
   'pnl_Mes.Caption = ""
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "CONTABILIZACION DE LA ASIGNACION DEL PBP: " & Trim(pnl_Periodo.Caption)
      .Range(.Cells(2, 2), .Cells(2, 7)).Merge
      .Range(.Cells(2, 2), .Cells(2, 7)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 7)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 2), .Cells(2, 7)).Font.Size = 12
      
      .Range(.Cells(2, 2), .Cells(2, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(2, 2), .Cells(2, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(2, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(2, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(2, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(2, 2), .Cells(2, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Cells(5, 2) = "N°"
      .Cells(5, 3) = "PRODUCTO"
      .Cells(5, 4) = "GLOSA"
      .Cells(5, 5) = "NRO. CUENTA"
      .Cells(5, 6) = "DEBE (S/.)"
      .Cells(5, 7) = "HABER (S/.)"
      .Range(.Cells(5, 2), .Cells(5, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(5, 2), .Cells(5, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
      .Range(.Cells(5, 2), .Cells(5, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(5, 2), .Cells(5, 7)).Font.Bold = True
      .Range(.Cells(5, 2), .Cells(5, 7)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 5
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 35
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 40
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 20
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 14
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 14
      .Columns("G").NumberFormat = "###,###,##0.00"
      
      .Range(.Cells(5, 1), .Cells(5, 7)).Font.Name = "Calibri"
      .Range(.Cells(5, 1), .Cells(5, 7)).Font.Size = 11
      
      Dim swtDebe As Double
      Dim swtHaber As Double
      
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
         
         .Range(.Cells(r_int_NumFil + 3, 2), .Cells(r_int_NumFil + 3, 7)).Font.Size = 10
         
         .Cells(r_int_NumFil + 3, 2) = r_int_NumFil - 2
         .Cells(r_int_NumFil + 3, 3) = "'" & CStr(pnl_NomProd)
         .Cells(r_int_NumFil + 3, 4) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 0)) 'Glosa
         .Cells(r_int_NumFil + 3, 5) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 3, 1)  'NroCuenta
         .Cells(r_int_NumFil + 3, 6) = grd_Listad.TextMatrix(r_int_NumFil - 3, 2) 'Debe
         .Cells(r_int_NumFil + 3, 7) = grd_Listad.TextMatrix(r_int_NumFil - 3, 3) 'Haber
         
         swtDebe = swtDebe + grd_Listad.TextMatrix(r_int_NumFil - 3, 2)
         swtHaber = swtHaber + grd_Listad.TextMatrix(r_int_NumFil - 3, 3)
         
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
      
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 7)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 7)).Font.Bold = True
      .Cells(r_int_NumFil + 3, 5).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NumFil + 3, 5) = "TOTAL : " & Space(5)
      .Cells(r_int_NumFil + 3, 6) = swtDebe
      .Cells(r_int_NumFil + 3, 7) = swtHaber
      
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(r_int_NumFil + 3, 5), .Cells(r_int_NumFil + 3, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Buscar_detalle()
Dim r_str_CodAux     As String
Dim r_str_Cadena     As String
Dim r_dbl_CapFMV     As Double
Dim r_dbl_CapMVPE    As Double
Dim r_dbl_CapMV      As Double
Dim r_str_MtoCof     As String
Dim r_str_CapCof     As String
Dim r_str_IntCof     As String
Dim r_str_CapCof2    As String
Dim r_str_CapCof3    As String
Dim r_str_Codigo     As String

   r_str_Codigo = moddat_g_str_CodPrd
   pnl_NomProd.Caption = Trim(moddat_g_str_NomPrd)
   pnl_Periodo.Caption = UCase(MesEspanol(moddat_g_int_EdaMes)) & " - " & moddat_g_int_EdaAno
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT B.HIPMAE_CODPRD, SUM(A.DETPBP_CAPCLI) AS CAPCLI, SUM(A.DETPBP_INTCLI) AS INTCLI, "
   g_str_Parame = g_str_Parame & "        SUM(A.DETPBP_CAPADE) AS CAPADE, SUM(A.DETPBP_INTADE) AS INTADE "
   g_str_Parame = g_str_Parame & " FROM   CRE_DETPBP A, CRE_HIPMAE B "
   g_str_Parame = g_str_Parame & " WHERE  DETPBP_PERMES = " & moddat_g_int_EdaMes & " "
   g_str_Parame = g_str_Parame & "   AND  DETPBP_PERANO = " & moddat_g_int_EdaAno & " "
   g_str_Parame = g_str_Parame & "   AND  DETPBP_NUMOPE = HIPMAE_NUMOPE "
   g_str_Parame = g_str_Parame & "   AND  DETPBP_FLGPBP = 1 "
   g_str_Parame = g_str_Parame & "   AND  HIPMAE_CODPRD IN (" & moddat_g_str_CodPrd & ") "
   g_str_Parame = g_str_Parame & " GROUP  BY HIPMAE_CODPRD "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do Until g_rst_Princi.EOF
      r_str_Codigo = g_rst_Princi!HIPMAE_CODPRD
      
      Select Case r_str_Codigo
         'CREDITO CRC-PBP
         Case "001"
            r_str_MtoCof = "152719010110"
            r_str_CapCof = "142104240101"
            r_str_IntCof = "152719010105"
            
         'CREDITO CME
         Case "003"
            r_str_MtoCof = "151719010103"
            r_str_CapCof = "141104250101"
            r_str_IntCof = "151719010105"
         
         'CREDITO PROYECTO MIHOGAR
         Case "004"
            r_str_MtoCof = "261202010101"
            r_str_CapCof = "141104230101"
                 
         'CREDITO MICASITA SOLES
         Case "006"
            r_str_MtoCof = "511401040601"
            r_str_CapCof = "141104060101"
            r_str_IntCof = "511401040601"
         
         'CREDITO MIVIVIENDA
         Case "007", "009", "010", "012", "013", "014", "015", "016", "017", "018", "019"
            r_str_MtoCof = "191807010101"
            r_str_CapCof = "141104230102"
            r_str_CapCof2 = "141104230104"
            r_str_CapCof3 = "141104230103"
      End Select
      
      Select Case r_str_Codigo
         Case "001", "003"
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP " & IIf(r_str_Codigo = "001", "CRC", "CME")
            grd_Listad.Col = 1: grd_Listad.Text = r_str_MtoCof
            grd_Listad.Col = 2: grd_Listad.Text = Format(g_rst_Princi!CAPADE + g_rst_Princi!INTADE, "###,###,##0.00")
            grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP " & IIf(r_str_Codigo = "001", "CRC", "CME") & " - CAP"
            grd_Listad.Col = 1: grd_Listad.Text = r_str_CapCof
            grd_Listad.Col = 2: grd_Listad.Text = Format(0, "###,###,##0.00")
            grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!CAPADE, "###,###,##0.00")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP " & IIf(r_str_Codigo = "001", "CRC", "CME") & " - INT"
            grd_Listad.Col = 1: grd_Listad.Text = r_str_IntCof
            grd_Listad.Col = 2: grd_Listad.Text = Format(0, "###,###,##0.00")
            grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!INTADE, "###,###,##0.00")
            
       Case "006"
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP MICASITA"
            grd_Listad.Col = 1: grd_Listad.Text = r_str_MtoCof
            grd_Listad.Col = 2: grd_Listad.Text = Format(g_rst_Princi!CAPCLI + g_rst_Princi!INTCLI, "###,###,##0.00")
            grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP MICASITA - CAP"
            grd_Listad.Col = 1: grd_Listad.Text = r_str_CapCof
            grd_Listad.Col = 2: grd_Listad.Text = Format(0, "###,###,##0.00")
            grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!CAPCLI, "###,###,##0.00")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP MICASITA - INT"
            grd_Listad.Col = 1: grd_Listad.Text = r_str_IntCof
            grd_Listad.Col = 2: grd_Listad.Text = Format(0, "###,###,##0.00")
            grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!INTCLI, "###,###,##0.00")
            
        Case "004"
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP MIHOGAR"
            grd_Listad.Col = 1: grd_Listad.Text = r_str_MtoCof
            grd_Listad.Col = 2: grd_Listad.Text = Format(g_rst_Princi!CAPADE, "###,###,##0.00")
            grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
            
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP MIHOGAR - CAP"
            grd_Listad.Col = 1: grd_Listad.Text = r_str_CapCof
            grd_Listad.Col = 2: grd_Listad.Text = Format(0, "###,###,##0.00")
            grd_Listad.Col = 3: grd_Listad.Text = Format(g_rst_Princi!CAPADE, "###,###,##0.00")
            
        Case "007", "009", "010", "013", "012", "014", "015", "016", "017", "018" '"019", "021", "022", "023"
            If r_str_Codigo = "009" Then ' 'swtCodProd
               r_dbl_CapFMV = r_dbl_CapFMV + g_rst_Princi!CAPADE
            ElseIf r_str_Codigo = "010" Then
               r_dbl_CapMVPE = r_dbl_CapMVPE + g_rst_Princi!CAPADE
            Else
               r_dbl_CapMV = r_dbl_CapMV + g_rst_Princi!CAPADE
            End If
            
      End Select
      
      'grd_Listad.Col = 0: grd_Listad.Text = Space(2) & grd_Listad.Text
      g_rst_Princi.MoveNext
   Loop
   
   If r_dbl_CapFMV > 0 Or r_dbl_CapMVPE > 0 Or r_dbl_CapMV > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP MIVIVIENDA"
      grd_Listad.Col = 1: grd_Listad.Text = r_str_MtoCof
      grd_Listad.Col = 2: grd_Listad.Text = Format(r_dbl_CapFMV + r_dbl_CapMVPE + r_dbl_CapMV, "###,###,##0.00")
      grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP MIVIVIENDA - CAP"
      grd_Listad.Col = 1: grd_Listad.Text = r_str_CapCof
      grd_Listad.Col = 2: grd_Listad.Text = Format(0, "###,###,##0.00")
      grd_Listad.Col = 3: grd_Listad.Text = Format(r_dbl_CapMV, "###,###,##0.00")
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP MIVIVIENDA - PER EXT CAP"
      grd_Listad.Col = 1: grd_Listad.Text = r_str_CapCof2
      grd_Listad.Col = 2: grd_Listad.Text = Format(0, "###,###,##0.00")
      grd_Listad.Col = 3: grd_Listad.Text = Format(r_dbl_CapMVPE, "###,###,##0.00")
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0: grd_Listad.Text = "APLICACION PBP MIVIVIENDA - UNI AND CAP"
      grd_Listad.Col = 1: grd_Listad.Text = r_str_CapCof3
      grd_Listad.Col = 2: grd_Listad.Text = Format(0, "###,###,##0.00")
      grd_Listad.Col = 3: grd_Listad.Text = Format(r_dbl_CapFMV, "###,###,##0.00")
      
      r_dbl_CapFMV = 0: r_dbl_CapMVPE = 0: r_dbl_CapMV = 0
   End If
   
   Call Sumar_Columnas
End Sub

Private Sub Sumar_Columnas()
Dim r_dbl_TotDebe As Double
Dim r_dbl_TotHaber As Double
Dim r_int_Fila As Integer
    
   For r_int_Fila = 0 To grd_Listad.Rows - 1
      r_dbl_TotDebe = r_dbl_TotDebe + grd_Listad.TextMatrix(r_int_Fila, 2)
      r_dbl_TotHaber = r_dbl_TotHaber + grd_Listad.TextMatrix(r_int_Fila, 3)
   Next r_int_Fila
    
   pnl_Debe.Caption = Format(r_dbl_TotDebe, "###,###,##0.00") & " "
   pnl_Haber.Caption = Format(r_dbl_TotHaber, "###,###,##0.00") & " "
End Sub

Function MesEspanol(Mes As Integer) As String
   Select Case Mes
      Case 1:  MesEspanol = "Enero"
      Case 2:  MesEspanol = "Febrero"
      Case 3:  MesEspanol = "Marzo"
      Case 4:  MesEspanol = "Abril"
      Case 5:  MesEspanol = "Mayo"
      Case 6:  MesEspanol = "Junio"
      Case 7:  MesEspanol = "Julio"
      Case 8:  MesEspanol = "Agosto"
      Case 9:  MesEspanol = "Setiembre"
      Case 10: MesEspanol = "Octubre"
      Case 11: MesEspanol = "Noviembre"
      Case 12: MesEspanol = "Diciembre"
   End Select
End Function
