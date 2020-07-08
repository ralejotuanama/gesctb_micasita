VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptSun_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   2685
   ClientLeft      =   1875
   ClientTop       =   6240
   ClientWidth     =   7200
   Icon            =   "GesCtb_frm_843.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _Version        =   65536
      _ExtentX        =   12726
      _ExtentY        =   5001
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
         Left            =   30
         TabIndex        =   1
         Top             =   60
         Width           =   7125
         _Version        =   65536
         _ExtentX        =   12568
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
            Height          =   300
            Left            =   630
            TabIndex        =   2
            Top             =   30
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "SUNAT"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   270
            Left            =   630
            TabIndex        =   3
            Top             =   300
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Libro de Inventarios y Balances"
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
            Picture         =   "GesCtb_frm_843.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   780
         Width           =   7125
         _Version        =   65536
         _ExtentX        =   12568
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
            Left            =   6510
            Picture         =   "GesCtb_frm_843.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_843.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_843.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1230
            Top             =   30
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1185
         Left            =   30
         TabIndex        =   8
         Top             =   1470
         Width           =   7125
         _Version        =   65536
         _ExtentX        =   12568
         _ExtentY        =   2090
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
         Begin VB.ComboBox cmb_CodMes 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   120
            Width           =   2500
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1890
            TabIndex        =   10
            Top             =   450
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
         Begin VB.Label Label5 
            Caption         =   "Año:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label3 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   150
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frm_RptSun_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_dbl_Evalua(30039)   As Double
Dim r_int_PerMes  As Integer
Dim r_int_PerAno  As Integer

Private Sub cmd_ExpExc_Click()

   If cmb_CodMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodMes)
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
   
   r_int_PerMes = CInt(cmb_CodMes.ItemData(cmb_CodMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   
  Call fs_GenExc
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   Call moddat_gs_Carga_LisIte_Combo(cmb_CodMes, 1, "033")
   Call gs_CentraForm(Me)
   Call fs_Limpia
      
   Screen.MousePointer = 0
   
End Sub

Private Sub fs_Limpia()

   r_int_PerMes = Month(date)
   r_int_PerAno = Year(date)
   
   If Month(date) = 12 Then
      r_int_PerMes = 1
      r_int_PerAno = Year(date) - 1
   Else
      r_int_PerMes = Month(date) - 1
      r_int_PerAno = Year(date)
   End If
 
   Call gs_BuscarCombo_Item(cmb_CodMes, r_int_PerMes)
   ipp_PerAno.Text = Format(r_int_PerAno, "0000")
   
End Sub

Private Sub fs_GenExc()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer

   Dim r_str_FecIni     As String
   Dim r_str_FecFin     As String
   Dim r_int_Contad     As Integer
   
   Dim r_str_CtaCtb     As String
   Dim r_str_CtaNta     As String
   Dim r_int_ConTem     As Integer
   
   Erase r_dbl_Evalua
   
   Screen.MousePointer = 11
      
   g_str_Parame = "SELECT HIPCIE_TDOCLI, HIPCIE_NDOCLI, DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_NOMBRE, HIPCIE_NUMOPE, HIPCIE_FECDES, HIPCIE_TIPMON, HIPCIE_TIPCAM, "
   g_str_Parame = g_str_Parame & "HIPCIE_PRVGEN, HIPCIE_PRVESP, HIPCIE_PRVCAM, HIPCIE_PRVCIC, HIPCIE_PRVADC, HIPCIE_CODPRD FROM CRE_HIPCIE H, CLI_DATGEN C WHERE  "
   g_str_Parame = g_str_Parame & "HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & r_int_PerMes & " AND HIPCIE_PERANO = " & r_int_PerAno & " "
   'g_str_Parame = g_str_Parame & "ORDER BY DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_FECDES ASC"
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No hay datos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
         
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   r_obj_Excel.Sheets(1).Name = "Cobranzas Dudosas "
   
   With r_obj_Excel.Sheets(1)
         
      .Cells(1, 1) = "FORMATO 3.6: ""LIBRO DE INVENTARIOS Y BALANCES - DETALLE DEL SALDO DE LA CUENTA 19 - PROVISIÓN PARA CUENTAS DE COBRANZAS DUDOSAS"""
      '.Cells(2, 1) = "(" & cmb_TipMon.Text & ")"
      .Cells(4, 1) = "EJERCICIO: "
      .Cells(4, 3) = r_int_PerAno & "-" & r_int_PerMes
      .Cells(5, 1) = "RUC: "
      .Cells(5, 3) = "20511904162"
      .Cells(6, 1) = "DENOMINACIÓN O RAZÓN SOCIAL: "
      .Cells(6, 3) = "EDPYME MICASITA S.A."

      .Range(.Cells(1, 1), .Cells(8, 6)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(1, 1), .Cells(8, 1)).Font.Bold = True
      
      .Range(.Cells(1, 1), .Cells(1, 6)).Merge
      '.Range(.Cells(2, 1), .Cells(2, 2)).Merge
      '.Range(.Cells(4, 1), .Cells(4, 2)).Merge
      '.Range(.Cells(5, 1), .Cells(5, 2)).Merge
      .Range(.Cells(6, 1), .Cells(6, 2)).Merge
      
      '.Range(.Cells(4, 3), .Cells(4, 5)).Merge
      '.Range(.Cells(5, 3), .Cells(5, 5)).Merge
      '.Range(.Cells(6, 3), .Cells(6, 5)).Merge
      
      .Cells(8, 1) = "INFORMACIÓN DE DEUDORES"
      .Cells(9, 1) = "DOCUMENTO DE IDENTIDAD"
      .Cells(10, 1) = "TIPO (TABLA 2)"
      .Cells(10, 2) = "NÚMERO"
      .Cells(9, 3) = "APELLIDOS Y NOMBRES"
      .Cells(8, 4) = "CUENTAS POR COBRAR PROVISIONADAS"
      .Cells(9, 4) = "NÚMERO DEL DOCUMENTO"
      .Cells(9, 5) = "FECHA DEL DOCUMENTO"
      .Cells(9, 6) = "MONTO (S/.)"
      .Cells(9, 7) = "MONTO (US$)"
       
      .Columns("A").ColumnWidth = 15
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      '.Columns("A").NumberFormat = "@"
      
      .Columns("B").ColumnWidth = 15
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("B").NumberFormat = "@"
      
      .Columns("C").ColumnWidth = 45
      '.Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
                  
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("E").NumberFormat = "@"
            
      .Columns("F").ColumnWidth = 15
      '.Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("F").NumberFormat = "###,###,##0.00"
      
      .Columns("G").ColumnWidth = 15
      '.Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("G").NumberFormat = "###,###,##0.00"

      .Range(.Cells(8, 1), .Cells(8, 3)).Merge
      .Range(.Cells(9, 1), .Cells(9, 2)).Merge
      .Range(.Cells(9, 3), .Cells(10, 3)).Merge
      .Range(.Cells(8, 4), .Cells(8, 7)).Merge
      .Range(.Cells(9, 4), .Cells(10, 4)).Merge
      .Range(.Cells(9, 5), .Cells(10, 5)).Merge
      .Range(.Cells(9, 6), .Cells(10, 6)).Merge
      .Range(.Cells(9, 7), .Cells(10, 7)).Merge
      
      .Range(.Cells(8, 1), .Cells(10, 7)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(8, 1), .Cells(10, 7)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(8, 1), .Cells(10, 7)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(8, 1), .Cells(10, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(8, 1), .Cells(10, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(8, 1), .Cells(10, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(8, 1), .Cells(10, 7)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(8, 1), .Cells(10, 7)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(8, 1), .Cells(10, 7)).WrapText = True
      .Range(.Cells(8, 1), .Cells(10, 7)).Font.Bold = True
      
                     
      g_rst_Princi.MoveFirst
      
      r_int_ConVer = 12
      r_int_ConTem = 0
      
      Do While Not g_rst_Princi.EOF
         
         If Trim(g_rst_Princi!HIPCIE_TDOCLI) = 1 Then
            .Cells(r_int_ConVer, 1) = "1"
         ElseIf Trim(g_rst_Princi!HIPCIE_TDOCLI) = 2 Then
            .Cells(r_int_ConVer, 1) = "4"
         ElseIf Trim(g_rst_Princi!HIPCIE_TDOCLI) = 3 Or Trim(g_rst_Princi!HIPCIE_TDOCLI) = 4 Then
            .Cells(r_int_ConVer, 1) = "0"
         ElseIf Trim(g_rst_Princi!HIPCIE_TDOCLI) = 5 Then
            .Cells(r_int_ConVer, 1) = "7"
         End If
                 
         .Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!HIPCIE_NDOCLI)
         .Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!DATGEN_APEPAT) & " " & Trim(g_rst_Princi!DATGEN_APEMAT) & " " & Trim(g_rst_Princi!DATGEN_NOMBRE)
         .Cells(r_int_ConVer, 4) = gf_Formato_NumOpe(Trim(g_rst_Princi!HIPCIE_NUMOPE))
         .Cells(r_int_ConVer, 5) = gf_FormatoFecha(Trim(g_rst_Princi!HIPCIE_FECDES))
         
         If Trim(g_rst_Princi!HIPCIE_TIPMON) = 1 Then
            .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVESP + g_rst_Princi!HIPCIE_PRVCAM + g_rst_Princi!HIPCIE_PRVCIC + g_rst_Princi!HIPCIE_PRVADC)
            .Cells(r_int_ConVer, 7) = "0"
            r_dbl_Evalua(0) = r_dbl_Evalua(0) + Trim(g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVESP + g_rst_Princi!HIPCIE_PRVCAM + g_rst_Princi!HIPCIE_PRVCIC + g_rst_Princi!HIPCIE_PRVADC)
         ElseIf Trim(g_rst_Princi!HIPCIE_TIPMON) = 2 Then
            .Cells(r_int_ConVer, 6) = Format((g_rst_Princi!HIPCIE_TIPCAM * g_rst_Princi!HIPCIE_PRVGEN) + (g_rst_Princi!HIPCIE_TIPCAM * g_rst_Princi!HIPCIE_PRVESP) + (g_rst_Princi!HIPCIE_TIPCAM * g_rst_Princi!HIPCIE_PRVCAM) + (g_rst_Princi!HIPCIE_TIPCAM * g_rst_Princi!HIPCIE_PRVCIC) + (g_rst_Princi!HIPCIE_TIPCAM * g_rst_Princi!HIPCIE_PRVADC), "###,###,##0.00")
            .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVESP + g_rst_Princi!HIPCIE_PRVCAM + g_rst_Princi!HIPCIE_PRVCIC + g_rst_Princi!HIPCIE_PRVADC)
            r_dbl_Evalua(0) = r_dbl_Evalua(0) + Format((g_rst_Princi!HIPCIE_TIPCAM * g_rst_Princi!HIPCIE_PRVGEN) + (g_rst_Princi!HIPCIE_TIPCAM * g_rst_Princi!HIPCIE_PRVESP) + (g_rst_Princi!HIPCIE_TIPCAM * g_rst_Princi!HIPCIE_PRVCAM) + (g_rst_Princi!HIPCIE_TIPCAM * g_rst_Princi!HIPCIE_PRVCIC) + (g_rst_Princi!HIPCIE_TIPCAM * g_rst_Princi!HIPCIE_PRVADC), "###,###,##0.00")
            r_dbl_Evalua(1) = r_dbl_Evalua(1) + Trim(g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVESP + g_rst_Princi!HIPCIE_PRVCAM + g_rst_Princi!HIPCIE_PRVCIC + g_rst_Princi!HIPCIE_PRVADC)
         End If
                                                            
         r_int_ConVer = r_int_ConVer + 1
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
            
      g_str_Parame = "SELECT COMCIE_TDOCLI, COMCIE_NDOCLI, DATGEN_RAZSOC, COMCIE_NUMOPE, COMCIE_FECDES, COMCIE_CODPRD, COMCIE_TIPMON, COMCIE_TIPCAM, "
      g_str_Parame = g_str_Parame & "COMCIE_PRVGEN, COMCIE_PRVESP, COMCIE_PRVCAM, COMCIE_PRVCIC, COMCIE_PRVADC FROM CRE_COMCIE C, EMP_DATGEN E "
      g_str_Parame = g_str_Parame & "WHERE DATGEN_EMPTDO = COMCIE_TDOCLI AND DATGEN_EMPNDO = COMCIE_NDOCLI AND "
      g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & r_int_PerMes & " AND COMCIE_PERANO = " & r_int_PerAno & " "
      g_str_Parame = g_str_Parame & "ORDER BY COMCIE_FECDES ASC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Do While Not g_rst_Princi.EOF
         
         If Trim(g_rst_Princi!COMCIE_TDOCLI) = 7 Or Trim(g_rst_Princi!COMCIE_TDOCLI) = 8 Then
            .Cells(r_int_ConVer, 1) = "6"
         End If
                 
         .Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!COMCIE_NDOCLI)
         .Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!DATGEN_RAZSOC)
         .Cells(r_int_ConVer, 4) = gf_Formato_NumOpe(Trim(g_rst_Princi!COMCIE_NUMOPE))
         .Cells(r_int_ConVer, 5) = gf_FormatoFecha(Trim(g_rst_Princi!COMCIE_FECDES))
         
         If Trim(g_rst_Princi!COMCIE_TIPMON) = 1 Then
            .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!COMCIE_PRVGEN + g_rst_Princi!COMCIE_PRVESP + g_rst_Princi!COMCIE_PRVCAM + g_rst_Princi!COMCIE_PRVCIC + g_rst_Princi!COMCIE_PRVADC)
            .Cells(r_int_ConVer, 7) = "0"
            r_dbl_Evalua(0) = r_dbl_Evalua(0) + Trim(g_rst_Princi!COMCIE_PRVGEN + g_rst_Princi!COMCIE_PRVESP + g_rst_Princi!COMCIE_PRVCAM + g_rst_Princi!COMCIE_PRVCIC + g_rst_Princi!COMCIE_PRVADC)
         ElseIf Trim(g_rst_Princi!COMCIE_TIPMON) = 2 Then
            .Cells(r_int_ConVer, 6) = Format((g_rst_Princi!COMCIE_TIPCAM * g_rst_Princi!COMCIE_PRVGEN) + (g_rst_Princi!COMCIE_TIPCAM * g_rst_Princi!COMCIE_PRVESP) + (g_rst_Princi!COMCIE_TIPCAM * g_rst_Princi!COMCIE_PRVCAM) + (g_rst_Princi!COMCIE_TIPCAM * g_rst_Princi!COMCIE_PRVCIC) + (g_rst_Princi!COMCIE_TIPCAM * g_rst_Princi!COMCIE_PRVADC), "###,###,##0.00")
            .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!COMCIE_PRVGEN + g_rst_Princi!COMCIE_PRVESP + g_rst_Princi!COMCIE_PRVCAM + g_rst_Princi!COMCIE_PRVCIC + g_rst_Princi!COMCIE_PRVADC)
            r_dbl_Evalua(0) = r_dbl_Evalua(0) + Format((g_rst_Princi!COMCIE_TIPCAM * g_rst_Princi!COMCIE_PRVGEN) + (g_rst_Princi!COMCIE_TIPCAM * g_rst_Princi!COMCIE_PRVESP) + (g_rst_Princi!COMCIE_TIPCAM * g_rst_Princi!COMCIE_PRVCAM) + (g_rst_Princi!COMCIE_TIPCAM * g_rst_Princi!COMCIE_PRVCIC) + (g_rst_Princi!COMCIE_TIPCAM * g_rst_Princi!COMCIE_PRVADC), "###,###,##0.00")
            r_dbl_Evalua(1) = r_dbl_Evalua(1) + Trim(g_rst_Princi!COMCIE_PRVGEN + g_rst_Princi!COMCIE_PRVESP + g_rst_Princi!COMCIE_PRVCAM + g_rst_Princi!COMCIE_PRVCIC + g_rst_Princi!COMCIE_PRVADC)
         End If
                                                            
         r_int_ConVer = r_int_ConVer + 1
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
                  
      .Cells(r_int_ConVer, 4).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_ConVer, 4), .Cells(r_int_ConVer, 7)).Font.Bold = True
      .Range(.Cells(r_int_ConVer, 4), .Cells(r_int_ConVer, 5)).Merge
      .Cells(r_int_ConVer, 4) = "MONTO TOTAL PROVISIONADO"
      .Cells(r_int_ConVer, 6) = r_dbl_Evalua(0)
      .Cells(r_int_ConVer, 7) = r_dbl_Evalua(1)
            
      .Range(.Cells(1, 1), .Cells(r_int_ConVer + 3, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_ConVer + 3, 99)).Font.Size = 9
   
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
End Sub


