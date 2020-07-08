VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   5430
   ClientTop       =   4455
   ClientWidth     =   5265
   Icon            =   "GesCtb_frm_809.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2685
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   4736
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
         TabIndex        =   6
         Top             =   60
         Width           =   5205
         _Version        =   65536
         _ExtentX        =   9181
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
            TabIndex        =   7
            Top             =   30
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte de Detalle de Cuentas x Pagar"
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
            TabIndex        =   8
            Top             =   315
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "COFIDE"
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
            Picture         =   "GesCtb_frm_809.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   780
         Width           =   5205
         _Version        =   65536
         _ExtentX        =   9181
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
            Left            =   4590
            Picture         =   "GesCtb_frm_809.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_809.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   2760
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
         Height          =   1155
         Left            =   30
         TabIndex        =   10
         Top             =   1470
         Width           =   5205
         _Version        =   65536
         _ExtentX        =   9181
         _ExtentY        =   2037
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
         Begin VB.ComboBox cmb_CodPrd 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2775
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1530
            TabIndex        =   2
            Top             =   750
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
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   810
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   450
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_PerMes           As String
Dim l_str_PerAno           As String

Private Sub cmd_ExpExc_Click()
   If cmb_CodPrd.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodPrd)
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
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   l_str_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   l_str_PerAno = ipp_PerAno.Text
         
   Call fs_GenExc(l_str_PerMes, l_str_PerAno)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_CodPrd)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno = Mid(date, 7, 4)
   
   cmb_CodPrd.AddItem "CREDITO CME"
   cmb_CodPrd.ItemData(cmb_CodPrd.NewIndex) = CInt(3)
   cmb_CodPrd.AddItem "CREDITO MIHOGAR"
   cmb_CodPrd.ItemData(cmb_CodPrd.NewIndex) = CInt(4)
   cmb_CodPrd.AddItem "CREDITO MIVIVIENDA"
   cmb_CodPrd.ItemData(cmb_CodPrd.NewIndex) = CInt(7)
End Sub

Private Sub fs_Limpia()
   cmb_PerMes.ListIndex = -1
End Sub

Private Sub fs_GenExc(ByVal p_PerMes As String, ByVal p_PerAno As String)
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_str_EmpSeg     As String
   
   g_str_Parame = "SELECT HIPCUO_NUMOPE, HIPMAE_CODPRD, HIPMAE_TDOCLI, HIPMAE_NDOCLI, HIPMAE_OPEMVI, DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_NOMBRE, HIPCUO_NUMCUO, HIPMAE_MONEDA, HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_COMCOF "
   g_str_Parame = g_str_Parame & "FROM CRE_HIPMAE A, CRE_HIPCUO B, CLI_DATGEN C "
   g_str_Parame = g_str_Parame & "WHERE HIPCUO_NUMOPE = HIPMAE_NUMOPE  AND "
   g_str_Parame = g_str_Parame & "HIPMAE_TDOCLI = DATGEN_TIPDOC  AND "
   g_str_Parame = g_str_Parame & "HIPMAE_NDOCLI = DATGEN_NUMDOC  AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2              AND "
      
   If cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex) = 3 Then
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 5              AND "
      g_str_Parame = g_str_Parame & "HIPMAE_CODPRD = '" & Format(cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex), "000") & "' AND "
   ElseIf cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex) = 4 Then
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 3              AND "
      g_str_Parame = g_str_Parame & "HIPMAE_CODPRD = '" & Format(cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex), "000") & "' AND "
   ElseIf cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex) = 7 Then
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 3              AND "
      g_str_Parame = g_str_Parame & "(HIPMAE_CODPRD = '" & Format(cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex), "000") & "' OR "
      g_str_Parame = g_str_Parame & "HIPMAE_CODPRD = '" & Format(cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex) + 2, "000") & "' OR "
      g_str_Parame = g_str_Parame & "HIPMAE_CODPRD = '" & Format(cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex) + 3, "000") & "') AND "
   End If
   
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & l_str_PerAno & Format(l_str_PerMes, "00") & "01 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & l_str_PerAno & Format(l_str_PerMes, "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
   g_str_Parame = g_str_Parame & " ORDER BY HIPMAE_NUMOPE ASC"
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
'      .Range(.Cells(1, 10), .Cells(2, 10)).HorizontalAlignment = xlHAlignRight
'      .Cells(1, 10) = "Dpto. de Tecnología e Informática"
'      .Cells(2, 10) = "Desarrollo de Sistemas"
'
'      .Range(.Cells(5, 5), .Cells(5, 5)).HorizontalAlignment = xlHAlignCenter
'      .Range(.Cells(5, 5), .Cells(5, 5)).Font.Bold = True
'      .Range(.Cells(5, 5), .Cells(5, 5)).Font.Underline = xlUnderlineStyleSingle
'      .Cells(5, 5) = "Evaluación Crediticia"
   
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "PRODUCTO"
      .Cells(1, 4) = "DOI CLIENTE"
      .Cells(1, 5) = "NRO. OPER. MIVIVIENDA"
      .Cells(1, 6) = "APELLIDO PATERNO"
      .Cells(1, 7) = "APELLIDO MATERNO"
      .Cells(1, 8) = "NOMBRE"
      .Cells(1, 9) = "NRO. CUOTA"
      .Cells(1, 10) = "TIPO MONEDA"
      .Cells(1, 11) = "CAPITAL"
      .Cells(1, 12) = "INTERES"
      .Cells(1, 13) = "COM. COFIDE"
       
      .Range(.Cells(1, 1), .Cells(1, 13)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 13)).Interior.Color = RGB(146, 208, 80)
      
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 4
      
      .Columns("B").ColumnWidth = 14
      .Columns("B").HorizontalAlignment = xlHAlignCenter
            
      .Columns("C").ColumnWidth = 39
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 18
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      '.Columns("D").NumberFormat = "@"
      
      .Columns("E").ColumnWidth = 20
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("E").NumberFormat = "@"
      
      .Columns("F").ColumnWidth = 20
      '.Columns("F").HorizontalAlignment = xlHAlignCenter
            
      .Columns("G").ColumnWidth = 20
      '.Columns("G").HorizontalAlignment = xlHAlignCenter
            
      .Columns("H").ColumnWidth = 18
      '.Columns("H").HorizontalAlignment = xlHAlignCenter
                  
      .Columns("I").ColumnWidth = 11
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      '.Columns("I").NumberFormat = "###,###,##0.00"
      
      .Columns("J").ColumnWidth = 20
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      '.Columns("J").NumberFormat = "###,###,##0.00"
      
      .Columns("K").ColumnWidth = 11
      '.Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("K").NumberFormat = "###,###,##0.00"
      
      .Columns("L").ColumnWidth = 11
      '.Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("L").NumberFormat = "###,###,##0.00"
      
      .Columns("M").ColumnWidth = 11
      '.Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("M").NumberFormat = "###,###,##0.00"
      
      g_rst_Princi.MoveFirst
      r_int_ConVer = 2
      Do While Not g_rst_Princi.EOF
      
         .Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         .Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!HIPCUO_NUMOPE)
         .Cells(r_int_ConVer, 3) = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
         .Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
         .Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!HIPMAE_OPEMVI)
         .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!DatGen_ApePat)
         .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!DatGen_ApeMat)
         .Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!DatGen_Nombre)
         .Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!HIPCUO_NUMCUO)
         .Cells(r_int_ConVer, 10) = IIf(g_rst_Princi!HIPMAE_MONEDA = 1, "SOLES", "DOLARES AMERICANOS")
         .Cells(r_int_ConVer, 11) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 12, 2)
         .Cells(r_int_ConVer, 12) = gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 12, 2)
         .Cells(r_int_ConVer, 13) = gf_FormatoNumero(g_rst_Princi!HIPCUO_COMCOF, 12, 2)
                                   
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 13)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 13)).Font.Size = 8
   End With
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

