VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Pro_CtbPrv_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   Icon            =   "GesCtb_frm_919.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel13 
      Height          =   8055
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9150
      _Version        =   65536
      _ExtentX        =   16140
      _ExtentY        =   14208
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
         Height          =   615
         Left            =   60
         TabIndex        =   3
         Top             =   60
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
            Height          =   465
            Left            =   630
            TabIndex        =   4
            Top             =   90
            Width           =   3765
            _Version        =   65536
            _ExtentX        =   6641
            _ExtentY        =   820
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
            Picture         =   "GesCtb_frm_919.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1365
         Left            =   60
         TabIndex        =   5
         Top             =   2280
         Width           =   9030
         _Version        =   65536
         _ExtentX        =   15928
         _ExtentY        =   2408
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
            Height          =   855
            Left            =   120
            TabIndex        =   6
            Top             =   380
            Width           =   8735
            _ExtentX        =   15399
            _ExtentY        =   1508
            _Version        =   393216
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_MesAc 
            Height          =   285
            Left            =   5760
            TabIndex        =   7
            Top             =   120
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mes Actual"
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
         Begin Threed.SSPanel pnl_glosa 
            Height          =   285
            Left            =   120
            TabIndex        =   8
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
         Begin Threed.SSPanel pnl_MesAnt 
            Height          =   285
            Left            =   4080
            TabIndex        =   9
            Top             =   120
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mes Anterior"
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
         Begin Threed.SSPanel pnl_Ajuste 
            Height          =   285
            Left            =   7200
            TabIndex        =   10
            Top             =   120
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Ajuste"
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   645
         Left            =   60
         TabIndex        =   11
         Top             =   710
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   8370
            Picture         =   "GesCtb_frm_919.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   45
            Picture         =   "GesCtb_frm_919.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   855
         Left            =   60
         TabIndex        =   12
         Top             =   1380
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
            TabIndex        =   13
            Top             =   120
            Width           =   2265
            _Version        =   65536
            _ExtentX        =   3995
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   4210752
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
         Begin Threed.SSPanel pnl_NomProd 
            Height          =   315
            Left            =   1320
            TabIndex        =   14
            Top             =   450
            Width           =   5985
            _Version        =   65536
            _ExtentX        =   10557
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   4210752
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Periodo:"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   165
            Width           =   585
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   525
            Width           =   555
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   4305
         Left            =   60
         TabIndex        =   17
         Top             =   3690
         Width           =   9030
         _Version        =   65536
         _ExtentX        =   15928
         _ExtentY        =   7594
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
         Begin Threed.SSPanel pnl_Total4 
            Height          =   315
            Left            =   6990
            TabIndex        =   18
            Top             =   3570
            Visible         =   0   'False
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
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
         Begin Threed.SSPanel pnl_Total3 
            Height          =   315
            Left            =   5250
            TabIndex        =   19
            Top             =   3570
            Visible         =   0   'False
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad2 
            Height          =   3165
            Left            =   90
            TabIndex        =   20
            Top             =   390
            Width           =   8850
            _ExtentX        =   15610
            _ExtentY        =   5583
            _Version        =   393216
            Rows            =   11
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Total2 
            Height          =   315
            Left            =   3450
            TabIndex        =   21
            Top             =   3570
            Visible         =   0   'False
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
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
         Begin Threed.SSPanel pnl_Total1 
            Height          =   315
            Left            =   1650
            TabIndex        =   22
            Top             =   3570
            Visible         =   0   'False
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
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
         Begin Threed.SSPanel pnl_Total5 
            Height          =   315
            Left            =   8730
            TabIndex        =   25
            Top             =   3570
            Visible         =   0   'False
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
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
         Begin VB.Label lbldetalle 
            Height          =   255
            Left            =   150
            TabIndex        =   24
            Top             =   3990
            Visible         =   0   'False
            Width           =   8775
         End
         Begin VB.Label Label8 
            Caption         =   "Detalle Periodo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   150
            Width           =   3885
         End
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
   grd_Listad.ColWidth(0) = 3945 '3975 'Glosa
   grd_Listad.ColWidth(1) = 1680 '2150 'Mes Anterior
   grd_Listad.ColWidth(2) = 1470 '1245 'Mes Actual
   grd_Listad.ColWidth(3) = 1480 '1245 'Ajuste
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   
   grd_Listad2.ColWidth(0) = 1525 'Operación
   grd_Listad2.ColWidth(1) = 1795 'Mes Diciembre
   grd_Listad2.ColWidth(2) = 1795 'Mes Actual
   grd_Listad2.ColWidth(3) = 1685 'Ingresos
   grd_Listad2.ColWidth(4) = 1685 'Gastos
   
   pnl_Total1.Caption = "0.00000"
   pnl_Total2.Caption = "0.00000"
   pnl_Total3.Caption = "0.00000"
   pnl_Total4.Caption = "0.00000"
   pnl_Total5.Caption = "0.00000"
   
   grd_Listad.Rows = 0
   grd_Listad2.Rows = 0
End Sub

Private Sub fs_Limpiar()
   pnl_NomProd.Caption = ""
   pnl_periodo.Caption = ""
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_Contad     As Integer
Dim r_int_NumFil     As Integer
Dim r_int_NumCol     As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "CONTABILIZACION DE PROVISIONES: " & Trim(pnl_periodo.Caption)
      .Range(.Cells(2, 2), .Cells(2, 7)).Merge
      .Range(.Cells(2, 2), .Cells(2, 7)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 7)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(4, 2) = "CUENTA"
      .Cells(4, 3) = pnl_MesAnt.Caption
      .Cells(4, 4) = pnl_MesAc.Caption
      .Cells(4, 5) = "AJUSTE"
      
      .Range(.Cells(4, 2), .Cells(4, 5)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 2), .Cells(4, 5)).Font.Bold = True
      .Range(.Cells(4, 3), .Cells(4, 5)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 45
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 20
      .Columns("C").HorizontalAlignment = xlHAlignRight
      .Columns("C").NumberFormat = "###,###,##0.00"
      .Columns("D").ColumnWidth = 20
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("E").ColumnWidth = 20
      .Columns("E").HorizontalAlignment = xlHAlignRight
      .Columns("E").NumberFormat = "###,###,##0.00"
      
      .Range(.Cells(1, 1), .Cells(10, 5)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 5)).Font.Size = 11
      .Range(.Cells(4, 2), .Cells(4, 5)).HorizontalAlignment = xlHAlignCenter
      
      r_int_NumFil = 5
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil, 2) = "'" & CStr(pnl_NomProd)
         .Cells(r_int_NumFil, 3) = grd_Listad.TextMatrix(0, 1)  'MES ANTERIOR
         .Cells(r_int_NumFil, 4) = grd_Listad.TextMatrix(0, 2)  'MES ACTUAL
         .Cells(r_int_NumFil, 5) = grd_Listad.TextMatrix(0, 3)  'AJUSTE
        
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
 
   End With
   
   If grd_Listad2.Rows > 0 Then
      With r_obj_Excel.ActiveSheet
         r_int_NumFil = r_int_NumFil + 1
         .Cells(r_int_NumFil, 2) = UCase(Me.Label8.Caption)
         .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, 7)).Merge
         .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, 7)).Font.Bold = True
         .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, 7)).HorizontalAlignment = xlHAlignCenter
         
         r_int_NumFil = r_int_NumFil + 2
         For r_int_Contad = 0 To grd_Listad2.Cols - 1
            .Cells(r_int_NumFil, r_int_Contad + 2) = CStr(grd_Listad2.TextMatrix(0, r_int_Contad))
            If r_int_Contad > 0 Then
               .Cells(r_int_NumFil, r_int_Contad + 2).ColumnWidth = 20
            End If
         Next r_int_Contad
          
         .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, grd_Listad2.Cols + 1)).Interior.Color = RGB(146, 208, 80)
         .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, grd_Listad2.Cols + 1)).Font.Bold = True
         .Range(.Cells(r_int_NumFil, 2), .Cells(r_int_NumFil, grd_Listad2.Cols + 1)).HorizontalAlignment = xlHAlignCenter
         
         r_int_NumFil = r_int_NumFil + 1
         For r_int_Contad = 1 To grd_Listad2.Rows - 1
            For r_int_NumCol = 0 To grd_Listad2.Cols - 1
               .Cells(r_int_NumFil, r_int_NumCol + 2) = IIf(r_int_NumCol = 0, "'" & CStr(grd_Listad2.TextMatrix(r_int_Contad, r_int_NumCol)), grd_Listad2.TextMatrix(r_int_Contad, r_int_NumCol))
            Next r_int_NumCol
            r_int_NumFil = r_int_NumFil + 1
         Next r_int_Contad
         .Range(.Cells(r_int_NumFil, 3), .Cells(r_int_NumFil, grd_Listad2.Cols + 1)).FormulaR1C1 = "=SUM(R[-" & r_int_NumFil - 8 & "]C:R[-1]C)"
         .Range(.Cells(r_int_NumFil, 3), .Cells(r_int_NumFil, grd_Listad2.Cols + 1)).Interior.Color = RGB(146, 208, 80)
         .Range(.Cells(r_int_NumFil, 3), .Cells(r_int_NumFil, grd_Listad2.Cols + 1)).Font.Bold = True
         .Range(.Cells(r_int_NumFil, 3), .Cells(r_int_NumFil, grd_Listad2.Cols + 1)).HorizontalAlignment = xlHAlignCenter
      
         If lbldetalle.Caption <> "" Then
            r_obj_Excel.Cells(r_int_NumFil + 2, 4) = lbldetalle.Caption
         End If
      
      End With
   End If
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Buscar_detalle()
Dim r_str_CtaAux        As String
Dim r_str_TipProv       As String
Dim r_str_TipMon        As String
Dim r_int_NumTipMon     As Integer
Dim r_dbl_MtoAnt        As Double
Dim r_dbl_MtoAct        As Double
Dim r_dbl_MtoAju        As Double

   r_str_TipMon = moddat_g_str_Moneda
   r_str_TipProv = Trim(moddat_g_str_Descri)
   r_dbl_MtoAnt = CDbl(moddat_g_dbl_MtoPre)
   r_dbl_MtoAct = CDbl(moddat_g_dbl_SalCap)
   r_dbl_MtoAju = CDbl(moddat_g_dbl_IngDec)
   r_str_CtaAux = moddat_g_str_DesMod
   
   pnl_NomProd.Caption = r_str_CtaAux & " - " & Trim(r_str_TipProv)
   pnl_periodo.Caption = UCase(MesEspanol(moddat_g_str_CodMes)) & " - " & moddat_g_str_CodAno
   Label8.Caption = "Detalle Periodo " & pnl_periodo.Caption
   pnl_MesAnt.Caption = UCase(MesEspanol(IIf(moddat_g_str_CodMes = 1, 12, CInt(moddat_g_str_CodMes) - 1)))
   pnl_MesAc.Caption = UCase(MesEspanol(moddat_g_str_CodMes))
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0: grd_Listad.Text = Trim(r_str_TipProv)
   grd_Listad.Col = 1: grd_Listad.Text = Format(r_dbl_MtoAnt, "###,###,##0.00")
   grd_Listad.Col = 2: grd_Listad.Text = Format(r_dbl_MtoAct, "###,###,##0.00")
   grd_Listad.Col = 3: grd_Listad.Text = Format(r_dbl_MtoAju, "###,###,##0.00")
   
   If r_dbl_MtoAct = 0 Then
      Exit Sub
   End If
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT RPT_CODIGO   NUM_OPERACION     , RPT_VALNUM01 PROV_GENERICA_DIC  , RPT_VALNUM02 PROV_ESPECIFICA_DIC, RPT_VALNUM03 PROV_PROCICLICA_DIC, "
   g_str_Parame = g_str_Parame & "        RPT_VALNUM04 PROV_GEN_RC_DIC   , RPT_VALNUM05 PROV_PROCIC_RC_DIC , RPT_VALNUM06 PROV_VOLUNTARIA_DIC, RPT_VALNUM07 PROV_GENERICA      , "
   g_str_Parame = g_str_Parame & "        RPT_VALNUM08 PROV_ESPECIFICA   , RPT_VALNUM09 PROV_PROCICLICA    , RPT_VALNUM10 PROV_GENERICA_RC   , RPT_VALNUM11 PROV_PROCIC_RC     , "
   g_str_Parame = g_str_Parame & "        RPT_VALNUM12 PROV_VOLUNTARIA   , "
   g_str_Parame = g_str_Parame & "        RPT_VALNUM13 INGRESOS_PROVGEN  , RPT_VALNUM14 INGRESOS_PROVESP   , RPT_VALNUM15 INGRESOS_PROVCIC   , "
   g_str_Parame = g_str_Parame & "        RPT_VALNUM16 INGRESOS_PROVGENRC, RPT_VALNUM17 INGRESOS_PROVCICRC , RPT_VALNUM18 INGRESOS_PROVOL    , RPT_VALNUM19 GASTOS_PROVGEN     , "
   g_str_Parame = g_str_Parame & "        RPT_VALNUM20 GASTOS_PROVESP    , RPT_VALNUM21 GASTOS_PROVCIC     , RPT_VALNUM22 GASTOS_PROVGENRC   , RPT_VALNUM23 GASTOS_PROVCICRC   , "
   g_str_Parame = g_str_Parame & "        RPT_VALNUM24 GASTOS_PROVOL     , RPT_MONEDA, "
   
   g_str_Parame = g_str_Parame & "        (SELECT DISTINCT RPT_VALNUM25 "
   g_str_Parame = g_str_Parame & "           FROM RPT_TABLA_TEMP "
   g_str_Parame = g_str_Parame & "          WHERE RPT_PERMES = " & moddat_g_str_CodMes & ""
   g_str_Parame = g_str_Parame & "            AND RPT_PERANO = " & moddat_g_str_CodAno & ""
   g_str_Parame = g_str_Parame & "            AND RPT_TERCRE = '" & modgen_g_str_NombPC & "'"
   g_str_Parame = g_str_Parame & "            AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "'"
   g_str_Parame = g_str_Parame & "            AND RPT_NOMBRE = 'REPORTE DE PROVISIONES' "
   g_str_Parame = g_str_Parame & "            AND RPT_VALNUM25 IS NOT NULL) TIPOCAMBIO "
               
   g_str_Parame = g_str_Parame & "   FROM RPT_TABLA_TEMP "
   g_str_Parame = g_str_Parame & "  WHERE RPT_PERMES = " & moddat_g_str_CodMes & ""
   g_str_Parame = g_str_Parame & "    AND RPT_PERANO = " & moddat_g_str_CodAno & ""
   g_str_Parame = g_str_Parame & "    AND RPT_TERCRE = '" & modgen_g_str_NombPC & "'"
   g_str_Parame = g_str_Parame & "    AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "'"
   g_str_Parame = g_str_Parame & "    AND RPT_NOMBRE = 'REPORTE DE PROVISIONES' "
   g_str_Parame = g_str_Parame & "    AND RPT_MONEDA = " & r_str_TipMon & ""
   g_str_Parame = g_str_Parame & "  ORDER BY RPT_CODIGO ASC"
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_Listad2.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   grd_Listad2.Rows = grd_Listad2.Rows + 2
   grd_Listad2.Row = grd_Listad2.Rows - 1
   grd_Listad2.FixedRows = 1
   
   'CABECERA
   grd_Listad2.Row = 0
   grd_Listad2.Col = 0:       grd_Listad2.Text = "OPERACION": grd_Listad2.CellAlignment = flexAlignCenterCenter
   
   If r_str_CtaAux = "431204010101" Or r_str_CtaAux = "432204010101" Then
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_ESP_DIC": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 2:       grd_Listad2.Text = "PROV_ESP_" & Left(UCase(MesEspanol(moddat_g_str_CodMes)), 3): grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 3:       grd_Listad2.Text = "INGRESOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 4:       grd_Listad2.Text = "GASTOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
   ElseIf r_str_CtaAux = "431204020101" Or r_str_CtaAux = "432204020101" Then
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_GEN_DIC": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 2:       grd_Listad2.Text = "PROV_GEN_" & Left(UCase(MesEspanol(moddat_g_str_CodMes)), 3): grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 3:       grd_Listad2.Text = "INGRESOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 4:       grd_Listad2.Text = "GASTOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
   ElseIf r_str_CtaAux = "431204020201" Or r_str_CtaAux = "432204020201" Then
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_PROCIC_DIC": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 2:       grd_Listad2.Text = "PROV_PROCIC_" & Left(UCase(MesEspanol(moddat_g_str_CodMes)), 3): grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 3:       grd_Listad2.Text = "INGRESOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 4:       grd_Listad2.Text = "GASTOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
   ElseIf r_str_CtaAux = "431204030101" Or r_str_CtaAux = "432204030101" Then
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_VOL_DIC": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 2:       grd_Listad2.Text = "PROV_VOL_" & Left(UCase(MesEspanol(moddat_g_str_CodMes)), 3): grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 3:       grd_Listad2.Text = "INGRESOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 4:       grd_Listad2.Text = "GASTOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
   ElseIf r_str_CtaAux = "431210020101" Then
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_GEN_RC_DIC": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 2:       grd_Listad2.Text = "PROV_GEN_RC_" & Left(UCase(MesEspanol(moddat_g_str_CodMes)), 3): grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 3:       grd_Listad2.Text = "INGRESOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 4:       grd_Listad2.Text = "GASTOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
   ElseIf r_str_CtaAux = "431210020201" Then
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROCIC_RC_DIC": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 2:       grd_Listad2.Text = "PROCIC_RC_" & Left(UCase(MesEspanol(moddat_g_str_CodMes)), 3): grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 3:       grd_Listad2.Text = "INGRESOS": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 4:       grd_Listad2.Text = "GASTOS": grd_Listad2.CellAlignment = flexAlignCenterCenter

   ElseIf r_str_CtaAux = "541401010101" Or r_str_CtaAux = "542401010101" Then
      grd_Listad2.Col = 1:       grd_Listad2.Text = "ING_PRGEN": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 2:       grd_Listad2.Text = "ING_PRCIC ": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 3:       grd_Listad2.Text = "ING_PRGEN_RC": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 4:       grd_Listad2.Text = "ING_PRCIC_RC": grd_Listad2.CellAlignment = flexAlignCenterCenter
      grd_Listad2.Col = 5:       grd_Listad2.Text = "ING_PRVOL": grd_Listad2.CellAlignment = flexAlignCenterCenter
   ElseIf r_str_CtaAux = "541401010102" Or r_str_CtaAux = "542401010102" Then
      grd_Listad2.Col = 1:       grd_Listad2.Text = "ING_PROV_ESP": grd_Listad2.CellAlignment = flexAlignCenterCenter
      
   ElseIf r_str_CtaAux = "141904010101" Or r_str_CtaAux = "142904010101" Then
      grd_Listad2.Cols = 2
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_ESPECIFICA": grd_Listad2.CellAlignment = flexAlignCenterCenter
   
   ElseIf r_str_CtaAux = "141904020101" Or r_str_CtaAux = "142904020101" Then
      grd_Listad2.Cols = 2
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_GENERICA": grd_Listad2.CellAlignment = flexAlignCenterCenter
   
   ElseIf r_str_CtaAux = "141904020201" Or r_str_CtaAux = "142904020201" Then
      grd_Listad2.Cols = 2
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_PROCICLICA": grd_Listad2.CellAlignment = flexAlignCenterCenter
   
   ElseIf r_str_CtaAux = "141904030101" Then
      grd_Listad2.Cols = 2
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_VOLUNTARIA": grd_Listad2.CellAlignment = flexAlignCenterCenter
   
   ElseIf r_str_CtaAux = "141910020101" Then
      grd_Listad2.Cols = 2
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_GENERICA_RC": grd_Listad2.CellAlignment = flexAlignCenterCenter
   
   ElseIf r_str_CtaAux = "141910020201" Then
      grd_Listad2.Cols = 2
      grd_Listad2.Col = 1:       grd_Listad2.Text = "PROV_PROCIC_RC": grd_Listad2.CellAlignment = flexAlignCenterCenter
   End If
   
   grd_Listad2.Rows = grd_Listad2.Rows - 1
   Do While Not g_rst_Princi.EOF
         
      grd_Listad2.Rows = grd_Listad2.Rows + 1
      grd_Listad2.Row = grd_Listad2.Rows - 1
               
      'NUM_OPERACION
      grd_Listad2.Col = 0
      grd_Listad2.Text = g_rst_Princi!NUM_OPERACION
      grd_Listad2.CellAlignment = flexAlignCenterCenter
      
      If r_str_CtaAux = "431204010101" Then
         grd_Listad2.Cols = 5
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_ESPECIFICA_DIC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(g_rst_Princi!PROV_ESPECIFICA, "###,###,###,##0.00")
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVESP, "###,###,###,##0.00")
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!GASTOS_PROVESP, "###,###,###,##0.00")
         
      ElseIf r_str_CtaAux = "431204020101" Then
         grd_Listad2.Cols = 5
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_GENERICA_DIC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(g_rst_Princi!PROV_GENERICA, "###,###,###,##0.00")
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVGEN, "###,###,###,##0.00")
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!GASTOS_PROVGEN, "###,###,###,##0.00")
         
      ElseIf r_str_CtaAux = "431204020201" Then
         grd_Listad2.Cols = 5
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_PROCICLICA_DIC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(g_rst_Princi!PROV_PROCICLICA, "###,###,###,##0.00")
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVCIC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!GASTOS_PROVCIC, "###,###,###,##0.00")
         
      ElseIf r_str_CtaAux = "431204030101" Then
         grd_Listad2.Cols = 5
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_VOLUNTARIA_DIC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(g_rst_Princi!PROV_VOLUNTARIA, "###,###,###,##0.00")
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVOL, "###,###,###,##0.00")
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!GASTOS_PROVOL, "###,###,###,##0.00")
         
      ElseIf r_str_CtaAux = "431210020101" Then
         grd_Listad2.Cols = 5
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_GEN_RC_DIC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(g_rst_Princi!PROV_GENERICA_RC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVGENRC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!GASTOS_PROVGENRC, "###,###,###,##0.00")
         
      ElseIf r_str_CtaAux = "431210020201" Then
         grd_Listad2.Cols = 5
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_PROCIC_RC_DIC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(g_rst_Princi!PROV_PROCIC_RC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVCICRC, "###,###,###,##0.00")
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!GASTOS_PROVCICRC, "###,###,###,##0.00")
         
      ElseIf r_str_CtaAux = "432204010101" Then
         grd_Listad2.Cols = 5
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_ESPECIFICA_DIC, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(IIf(IsNull(g_rst_Princi!PROV_ESPECIFICA), 0, g_rst_Princi!PROV_ESPECIFICA), "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVESP, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!GASTOS_PROVESP, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
      ElseIf r_str_CtaAux = "432204020101" Then
         grd_Listad2.Cols = 5
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_GENERICA_DIC, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(IIf(IsNull(g_rst_Princi!PROV_GENERICA), 0, g_rst_Princi!PROV_GENERICA), "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVGEN, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!GASTOS_PROVGEN, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
      ElseIf r_str_CtaAux = "432204020201" Then
         grd_Listad2.Cols = 5
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_PROCICLICA_DIC, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(IIf(IsNull(g_rst_Princi!PROV_PROCICLICA), 0, g_rst_Princi!PROV_PROCICLICA), "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVCIC, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!GASTOS_PROVCIC, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
      ElseIf r_str_CtaAux = "432204030101" Then
         grd_Listad2.Cols = 5
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_VOLUNTARIA_DIC, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(g_rst_Princi!PROV_VOLUNTARIA, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVOL, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!GASTOS_PROVOL, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
      
      ElseIf r_str_CtaAux = "541401010101" And r_str_TipProv = "Revers.Provis.x Cred Direc" Then
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVGEN, "###,###,###,##0.00")
         grd_Listad2.ColWidth(1) = 1400
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVCIC, "###,###,###,##0.00")
         grd_Listad2.ColWidth(2) = 1400
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVGENRC, "###,###,###,##0.00")
         grd_Listad2.ColWidth(3) = 1400
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVCICRC, "###,###,###,##0.00")
         grd_Listad2.ColWidth(4) = 1400
          
         grd_Listad2.Col = 5
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVOL, "###,###,###,##0.00")
         grd_Listad2.ColWidth(5) = 1400
         
      ElseIf r_str_CtaAux = "541401010101" And r_str_TipProv = "Hipotecario MN Especifica" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVESP, "###,###,###,##0.00")
      
      ElseIf r_str_CtaAux = "542401010101" And r_str_TipProv = "Hipotecario ME Generica total" Then
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVGEN, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         grd_Listad2.ColWidth(1) = 1400
         
         grd_Listad2.Col = 2
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVCIC, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         grd_Listad2.ColWidth(2) = 1400
         
         grd_Listad2.Col = 3
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVGENRC, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         grd_Listad2.ColWidth(3) = 1400
         
         grd_Listad2.Col = 4
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVCICRC, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         grd_Listad2.ColWidth(4) = 1400
          
         grd_Listad2.Col = 5
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVOL, "###,###,###,##0.00") * g_rst_Princi!TIPOCAMBIO
         grd_Listad2.ColWidth(5) = 1400
         
      ElseIf r_str_CtaAux = "542401010101" And r_str_TipProv = "Hipotecario ME Especifica" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!INGRESOS_PROVESP, "###,###,###,##0.00") * IIf(IsNull(g_rst_Princi!TIPOCAMBIO), 0, g_rst_Princi!TIPOCAMBIO)
         
      ElseIf r_str_CtaAux = "141904010101" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_ESPECIFICA, "###,###,###,##0.00")
      ElseIf r_str_CtaAux = "141904020101" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_GENERICA, "###,###,###,##0.00")
      ElseIf r_str_CtaAux = "141904020201" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_PROCICLICA, "###,###,###,##0.00")
      ElseIf r_str_CtaAux = "141904030101" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_VOLUNTARIA, "###,###,###,##0.00")
      ElseIf r_str_CtaAux = "141910020101" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_GENERICA_RC, "###,###,###,##0.00")
      ElseIf r_str_CtaAux = "141910020201" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_PROCIC_RC, "###,###,###,##0.00")
      ElseIf r_str_CtaAux = "142904010101" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_ESPECIFICA, "###,###,###,##0.00")
      ElseIf r_str_CtaAux = "142904020101" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_GENERICA, "###,###,###,##0.00")
      ElseIf r_str_CtaAux = "142904020201" Then
         grd_Listad2.Cols = 2
         grd_Listad2.Col = 1
         grd_Listad2.Text = Format(g_rst_Princi!PROV_PROCICLICA, "###,###,###,##0.00")
      End If
         
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad2.Redraw = True
   If grd_Listad2.Rows > 0 Then
      grd_Listad2.Enabled = True
   End If
           
   Call gs_UbiIniGrid(grd_Listad2)
   Call gs_SetFocus(grd_Listad2)
   
   Call Sumar_Columnas

End Sub

Private Sub Sumar_Columnas()
Dim r_dbl_TotMesDic     As Double
Dim r_dbl_TotMesAct     As Double
Dim r_dbl_TotIngresos   As Double
Dim r_dbl_TotGastos     As Double
Dim r_dbl_TotOtros      As Double
Dim r_int_fila          As Integer
Dim r_dbl_TotAux        As Double
   
   pnl_Total1.Visible = True
   
   For r_int_fila = 1 To grd_Listad2.Rows - 1
      r_dbl_TotMesDic = r_dbl_TotMesDic + IIf(grd_Listad2.TextMatrix(r_int_fila, 1) = "", 0, grd_Listad2.TextMatrix(r_int_fila, 1))
      'pnl_Total1.Visible = True
      If grd_Listad2.Cols > 2 Then
         r_dbl_TotMesAct = r_dbl_TotMesAct + IIf(grd_Listad2.TextMatrix(r_int_fila, 2) = "", 0, grd_Listad2.TextMatrix(r_int_fila, 2))
         'pnl_Total2.Visible = True
         r_dbl_TotIngresos = r_dbl_TotIngresos + IIf(grd_Listad2.TextMatrix(r_int_fila, 3) = "", 0, grd_Listad2.TextMatrix(r_int_fila, 3))
         'pnl_Total3.Visible = True
         r_dbl_TotGastos = r_dbl_TotGastos + IIf(grd_Listad2.TextMatrix(r_int_fila, 4) = "", 0, grd_Listad2.TextMatrix(r_int_fila, 4))
         'pnl_Total4.Visible = True
      End If
      If grd_Listad2.Cols = 6 Then
         r_dbl_TotOtros = r_dbl_TotOtros + IIf(grd_Listad2.TextMatrix(r_int_fila, 5) = "", 0, grd_Listad2.TextMatrix(r_int_fila, 5))
      End If
   Next r_int_fila
   pnl_Total1.Caption = Format(r_dbl_TotMesDic, "###,###,##0.00")
   pnl_Total2.Caption = Format(r_dbl_TotMesAct, "###,###,##0.00")
   pnl_Total3.Caption = Format(r_dbl_TotIngresos, "###,###,##0.00")
   pnl_Total4.Caption = Format(r_dbl_TotGastos, "###,###,##0.00")
   pnl_Total5.Caption = Format(r_dbl_TotOtros, "###,###,##0.00")
   If grd_Listad2.Cols > 2 Then
      pnl_Total2.Visible = True
      pnl_Total3.Visible = True
      pnl_Total4.Visible = True
   End If
   If grd_Listad2.Cols = 6 Then
      pnl_Total1.Width = 1400
      pnl_Total2.Width = 1400
      pnl_Total3.Width = 1400
      pnl_Total4.Width = 1400
      pnl_Total5.Width = 1400
      
      pnl_Total2.Left = 3060
      pnl_Total3.Left = 4470
      pnl_Total4.Left = 5880
      pnl_Total5.Left = 7290
      
      pnl_Total5.Visible = True
      lbldetalle.Visible = True
      r_dbl_TotAux = -pnl_Total1.Caption - pnl_Total2.Caption - pnl_Total3.Caption - pnl_Total4.Caption - r_dbl_TotOtros
      lbldetalle.Caption = Trim(moddat_g_str_Descri) & " : - " & pnl_Total1.Caption & " - " & pnl_Total2.Caption & " - " & pnl_Total3.Caption & " - " & pnl_Total4.Caption & " - " & r_dbl_TotOtros & " = " & Format(r_dbl_TotAux, "###,###,##0.00")
   End If
   If pnl_NomProd = "542401010101 - Hipotecario ME Especifica" Then
      lbldetalle.Visible = True
      r_dbl_TotAux = -pnl_Total1.Caption - grd_Listad.TextMatrix(0, 2)
      lbldetalle.Caption = Trim(moddat_g_str_Descri) & " : " & pnl_Total1.Caption & " " & r_dbl_TotAux & " = " & (pnl_Total1.Caption + r_dbl_TotAux)
   End If
   
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
