VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_CtbBbp_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11115
   Icon            =   "GesCtb_frm_924.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11115
      _Version        =   65536
      _ExtentX        =   19606
      _ExtentY        =   12303
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
         TabIndex        =   4
         Top             =   60
         Width           =   10995
         _Version        =   65536
         _ExtentX        =   19394
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
            Left            =   570
            TabIndex        =   5
            Top             =   30
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Proceso"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   375
            Left            =   570
            TabIndex        =   6
            Top             =   270
            Width           =   5235
            _Version        =   65536
            _ExtentX        =   9234
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Contabilizaci�n de Desembolsos BBP"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.24
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
            Picture         =   "GesCtb_frm_924.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   645
         Left            =   60
         TabIndex        =   7
         Top             =   780
         Width           =   10995
         _Version        =   65536
         _ExtentX        =   19394
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
            Left            =   10410
            Picture         =   "GesCtb_frm_924.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Salir de la Opci�n"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   45
            Picture         =   "GesCtb_frm_924.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   510
         Left            =   60
         TabIndex        =   8
         Top             =   1470
         Width           =   10995
         _Version        =   65536
         _ExtentX        =   19394
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
         Begin Threed.SSPanel pnl_FecCar 
            Height          =   345
            Left            =   9060
            TabIndex        =   9
            Top             =   90
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   609
            _StockProps     =   15
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
         End
         Begin Threed.SSPanel pnl_NumCar 
            Height          =   345
            Left            =   2100
            TabIndex        =   10
            Top             =   90
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   609
            _StockProps     =   15
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
         End
         Begin VB.Label Label3 
            Caption         =   "N�mero de Carta FMV"
            Height          =   255
            Left            =   135
            TabIndex        =   12
            Top             =   165
            Width           =   1875
         End
         Begin VB.Label lbl_FecCar 
            Caption         =   "Fecha Carta FMV:"
            Height          =   255
            Left            =   7530
            TabIndex        =   11
            Top             =   165
            Width           =   1575
         End
      End
      Begin Threed.SSPanel SSPanel22 
         Height          =   4875
         Left            =   60
         TabIndex        =   13
         Top             =   2025
         Width           =   10995
         _Version        =   65536
         _ExtentX        =   19394
         _ExtentY        =   8599
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
            TabIndex        =   14
            Top             =   60
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operacion"
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
            Left            =   1590
            TabIndex        =   15
            Top             =   60
            Width           =   5100
            _Version        =   65536
            _ExtentX        =   8996
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre del Cliente"
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
         Begin Threed.SSPanel pnl_Tit_MtoDes 
            Height          =   285
            Left            =   6697
            TabIndex        =   16
            Top             =   60
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Desembolso"
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
            Height          =   4425
            Left            =   60
            TabIndex        =   3
            Top             =   390
            Width           =   10875
            _ExtentX        =   19182
            _ExtentY        =   7805
            _Version        =   393216
            Rows            =   18
            Cols            =   11
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_CtaCtb 
            Height          =   285
            Left            =   7897
            TabIndex        =   17
            Top             =   60
            Width           =   1620
            _Version        =   65536
            _ExtentX        =   2857
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cuenta Contable"
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
         Begin Threed.SSPanel pnl_Tit_TipCta 
            Height          =   285
            Left            =   9517
            TabIndex        =   18
            Top             =   60
            Width           =   930
            _Version        =   65536
            _ExtentX        =   1640
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Cta."
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
Attribute VB_Name = "frm_Pro_CtbBbp_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("�Est� seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
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
   
   Call fs_Limpia
   Call fs_BuscaDatos
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(grd_Listad)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
   pnl_NumCar.Caption = ""
   pnl_FecCar.Caption = ""
   grd_Listad.Cols = 5
   grd_Listad.ColWidth(0) = 1500
   grd_Listad.ColWidth(1) = 5100
   grd_Listad.ColWidth(2) = 1200
   grd_Listad.ColWidth(3) = 1620
   grd_Listad.ColWidth(4) = 930
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_BuscaDatos()
Dim r_dbl_ComDes     As Double
Dim r_dbl_SumDes     As Double
Dim r_dbl_SumPre     As Double
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(EVACOF_NUMSOL)       AS NRO_SOLICITUD, "
   g_str_Parame = g_str_Parame & "       TRIM(EVACOF_FECREC)       AS FECHA_RECEPCION, "
   g_str_Parame = g_str_Parame & "       TRIM(SOLMAE_TITTDO)||'-'||TRIM(SOLMAE_TITNDO) AS DOCUMENTO_CLIENTE, "
   g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
   g_str_Parame = g_str_Parame & "       B.SOLMAE_FMVBBP           AS MONTO_DESEMBOLSO, "
   g_str_Parame = g_str_Parame & "       NVL(D.HIPMAE_NUMOPE, '-') AS NRO_OPERACION "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVACOF A"
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE B ON B.SOLMAE_NUMERO = A.EVACOF_NUMSOL "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.SOLMAE_TITTDO AND C.DATGEN_NUMDOC = B.SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE D ON D.HIPMAE_NUMSOL = B.SOLMAE_NUMERO AND HIPMAE_SITUAC = 2 "
   g_str_Parame = g_str_Parame & " WHERE TRIM(A.EVACOF_NUMCAR) = '" & Trim(moddat_g_str_NumOpe) & "' "
   g_str_Parame = g_str_Parame & "   AND EVACOF_FECDES = " & Trim(moddat_g_str_FecDes) & " "
   g_str_Parame = g_str_Parame & " ORDER BY NOMBRE_CLIENTE "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      MsgBox "No se encontraron datos para la b�squeda.", vbInformation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Call fs_Limpia
      Exit Sub
   End If
   
   r_dbl_ComDes = 0
   r_dbl_SumDes = 0
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   pnl_NumCar.Caption = Trim(moddat_g_str_NumOpe)
   pnl_FecCar.Caption = gf_FormatoFecha(g_rst_Princi!FECHA_RECEPCION)
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = gf_Formato_NumOpe(g_rst_Princi!NRO_OPERACION)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Mid(Trim(g_rst_Princi!NRO_OPERACION) & " - " & Trim(g_rst_Princi!NOMBRE_CLIENTE), 1, 50)
        
      grd_Listad.Col = 2
      grd_Listad.Text = Format(CDbl(g_rst_Princi!MONTO_DESEMBOLSO), "###,###,##0.00")
      
      grd_Listad.Col = 3
      grd_Listad.Text = "291807010113"
      
      grd_Listad.Col = 4
      grd_Listad.Text = "H"
      

      r_dbl_SumDes = r_dbl_SumDes + Format(CDbl(g_rst_Princi!MONTO_DESEMBOLSO), "###,###,##0.00")

      g_rst_Princi.MoveNext
   Loop
   
   'Desembolso BBP
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0: grd_Listad.Text = "-"
   grd_Listad.Col = 1: grd_Listad.Text = "AB/ Fondo Mivivienda/BBP" & " - " & Trim(moddat_g_str_NumOpe)
   grd_Listad.Col = 2: grd_Listad.Text = Format(r_dbl_SumDes, "###,###,##0.00")
   grd_Listad.Col = 3: grd_Listad.Text = "111301060102"
   grd_Listad.Col = 4: grd_Listad.Text = "D"
     
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 2) = "CONTABILIZACION DE DESEMBOLSOS BBP"
      .Range(.Cells(1, 2), .Cells(1, 6)).Merge
      .Range(.Cells(1, 2), .Cells(1, 6)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(1, 6)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(3, 2) = "CARTA BBP:   " & Trim(pnl_NumCar.Caption) & "      -      " & "FECHA DE CARTA:   " & Trim(pnl_FecCar.Caption)
      .Range(.Cells(3, 2), .Cells(3, 6)).Merge
      .Range(.Cells(3, 2), .Cells(3, 6)).Font.Bold = True
      .Range(.Cells(3, 2), .Cells(3, 6)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(5, 2) = "NRO. OPERACION"
      .Cells(5, 3) = "NOMBRE DEL CLIENTE"
      .Cells(5, 4) = "DESEMBOLSO S/."
      .Cells(5, 5) = "CUENTA CONTABLE"
      .Cells(5, 6) = "TIPO CUENTA"
      
      .Range(.Cells(5, 2), .Cells(5, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(5, 2), .Cells(5, 6)).Font.Bold = True
      .Range(.Cells(5, 3), .Cells(5, 6)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 55
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 16
      .Columns("D").NumberFormat = "###,###,##0.00"
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("E").ColumnWidth = 20
      .Columns("E").NumberFormat = "@"
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 12
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(50, 6)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(50, 6)).Font.Size = 11
      
      r_int_NumFil = 3
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil + 3, 2) = grd_Listad.TextMatrix(r_int_NumFil - 3, 0)
         .Cells(r_int_NumFil + 3, 3) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 3, 1)
         .Cells(r_int_NumFil + 3, 4) = grd_Listad.TextMatrix(r_int_NumFil - 3, 2)
         .Cells(r_int_NumFil + 3, 5) = grd_Listad.TextMatrix(r_int_NumFil - 3, 3)
         .Cells(r_int_NumFil + 3, 6) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 3, 4)
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub
