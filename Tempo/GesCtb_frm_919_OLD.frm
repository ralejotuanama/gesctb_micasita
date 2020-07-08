VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Pro_CtbPrv_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000
   Icon            =   "GesCtb_frm_919.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel6 
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
         Picture         =   "GesCtb_frm_919.frx":000C
         Top             =   120
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   3750
      Left            =   0
      TabIndex        =   2
      Top             =   2145
      Width           =   9030
      _Version        =   65536
      _ExtentX        =   15928
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
         Left            =   7415
         TabIndex        =   3
         Top             =   3390
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
         Left            =   6115
         TabIndex        =   4
         Top             =   3390
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
      Begin Threed.SSPanel pnl_Tit_TipPpg 
         Height          =   285
         Left            =   4050
         TabIndex        =   5
         Top             =   120
         Width           =   2145
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
         Left            =   7445
         TabIndex        =   6
         Top             =   120
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
         Left            =   80
         TabIndex        =   7
         Top             =   120
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
         Left            =   6190
         TabIndex        =   8
         Top             =   120
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
   Begin Threed.SSPanel SSPanel10 
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
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   45
         Picture         =   "GesCtb_frm_919.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Exportar a Excel"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   8370
         Picture         =   "GesCtb_frm_919.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel2 
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
Attribute VB_Name = "frm_Pro_CtbPrv_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
   
   pnl_Debe.Caption = "0.00000"
   pnl_Haber.Caption = "0.00000"
   grd_Listad.Rows = 0
End Sub

Private Sub fs_Limpiar()
   pnl_NomProd.Caption = ""
   pnl_periodo.Caption = ""
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "CONTABILIZACION DE PROVISIONES: " & Trim(pnl_periodo.Caption)
      .Range(.Cells(2, 2), .Cells(2, 7)).Merge
      .Range(.Cells(2, 2), .Cells(2, 7)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 7)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(5, 2) = "TIPO DE PROVISIÓN"
      .Cells(5, 3) = "GLOSA"
      .Cells(5, 4) = "Nro CUENTA"
      .Cells(5, 5) = "DEBE"
      .Cells(5, 6) = "HABER"
      
      .Range(.Cells(5, 2), .Cells(5, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(5, 2), .Cells(5, 6)).Font.Bold = True
      .Range(.Cells(5, 3), .Cells(5, 6)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 35 '20
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
         .Cells(r_int_NumFil + 3, 2) = "'" & CStr(pnl_NomProd)
         .Cells(r_int_NumFil + 3, 3) = CStr(grd_Listad.TextMatrix(r_int_NumFil - 3, 0)) 'Glosa
         .Cells(r_int_NumFil + 3, 4) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 3, 1) 'NroCuenta
         .Cells(r_int_NumFil + 3, 5) = grd_Listad.TextMatrix(r_int_NumFil - 3, 2) 'Debe
         .Cells(r_int_NumFil + 3, 6) = grd_Listad.TextMatrix(r_int_NumFil - 3, 3) 'Haber
         
         swtDebe = swtDebe + grd_Listad.TextMatrix(r_int_NumFil - 3, 2)
         swtHaber = swtHaber + grd_Listad.TextMatrix(r_int_NumFil - 3, 3)
         
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

Private Sub fs_Buscar_detalle()
Dim r_str_CtaD          As String
Dim r_str_CtaH          As String
Dim r_str_CtaAux        As String
Dim r_str_TipProv       As String
Dim r_str_TipMon        As String
Dim r_int_NumTipMon     As Integer
Dim r_dbl_MtoDif        As Double

   r_str_TipMon = moddat_g_str_Moneda
   r_str_TipProv = Trim(moddat_g_str_Descri)
   r_dbl_MtoDif = CDbl(moddat_g_dbl_MtoPre)
   pnl_NomProd.Caption = Trim(r_str_TipProv)
   pnl_periodo.Caption = UCase(MesEspanol(moddat_g_int_EdaMes)) & " - " & moddat_g_int_EdaAno
   
   If r_str_TipMon = "DOLARES AMERICANOS" Then r_int_NumTipMon = 2 Else r_int_NumTipMon = 1
   
   If r_str_TipProv = "PROV. GENERICA" Then
      r_str_CtaD = "43" & r_int_NumTipMon & "204020101"
      r_str_CtaH = "14" & r_int_NumTipMon & "904020101"
   ElseIf r_str_TipProv = "PROV. ESPECIFICA" Then
      r_str_CtaD = "43" & r_int_NumTipMon & "204010101"
      r_str_CtaH = "14" & r_int_NumTipMon & "904010101"
   ElseIf r_str_TipProv = "PROV. PROCICLICA" Then
      r_str_CtaD = "43" & r_int_NumTipMon & "204020201"
      r_str_CtaH = "14" & r_int_NumTipMon & "904020201"
   ElseIf r_str_TipProv = "PROV. GENERICA RC" Then
      r_str_CtaD = "43" & r_int_NumTipMon & "210020101"
      r_str_CtaH = "14" & r_int_NumTipMon & "910020101"
   ElseIf r_str_TipProv = "PROV. PROCICLICA RC" Then
      r_str_CtaD = "43" & r_int_NumTipMon & "204020201"
      r_str_CtaH = "14" & r_int_NumTipMon & "910020201"
   ElseIf r_str_TipProv = "PROV. RIESGO CAMB." Then
      r_str_CtaD = "43" & r_int_NumTipMon & "204020201"
      r_str_CtaH = "14" & r_int_NumTipMon & "904020201"
   ElseIf r_str_TipProv = "PROV. VOLUNTARIA" Then
      r_str_CtaD = "43" & r_int_NumTipMon & "204030101"
      r_str_CtaH = "14" & r_int_NumTipMon & "904030101"
   End If
   
   If r_dbl_MtoDif < 0 Then r_str_CtaAux = r_str_CtaD: r_str_CtaD = r_str_CtaH: r_str_CtaH = r_str_CtaAux
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0: grd_Listad.Text = Trim(r_str_TipProv)
   grd_Listad.Col = 1: grd_Listad.Text = r_str_CtaD
   grd_Listad.Col = 2: grd_Listad.Text = Format(IIf(r_dbl_MtoDif < 0, r_dbl_MtoDif * -1, r_dbl_MtoDif), "###,###,##0.00")
   grd_Listad.Col = 3: grd_Listad.Text = Format(0, "###,###,##0.00")
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0: grd_Listad.Text = Trim(r_str_TipProv)
   grd_Listad.Col = 1: grd_Listad.Text = r_str_CtaH
   grd_Listad.Col = 2: grd_Listad.Text = Format(0, "###,###,##0.00")
   grd_Listad.Col = 3: grd_Listad.Text = Format(IIf(r_dbl_MtoDif < 0, r_dbl_MtoDif * -1, r_dbl_MtoDif), "###,###,##0.00")

   Call Sumar_Columnas
End Sub

Private Sub Sumar_Columnas()
Dim r_dbl_TotDebe       As Double
Dim r_dbl_TotHaber      As Double
Dim r_int_Fila          As Integer
    
    For r_int_Fila = 0 To grd_Listad.Rows - 1
        r_dbl_TotDebe = r_dbl_TotDebe + grd_Listad.TextMatrix(r_int_Fila, 2)
        r_dbl_TotHaber = r_dbl_TotHaber + grd_Listad.TextMatrix(r_int_Fila, 3)
    Next r_int_Fila
    
    pnl_Debe.Caption = Format(r_dbl_TotDebe, "###,###,##0.00")
    pnl_Haber.Caption = Format(r_dbl_TotHaber, "###,###,##0.00")
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
