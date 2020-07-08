VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_17 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2340
   ClientLeft      =   9990
   ClientTop       =   6285
   ClientWidth     =   5445
   Icon            =   "GesCtb_frm_815.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   4207
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
         Left            =   30
         TabIndex        =   1
         Top             =   60
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
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
            Left            =   570
            TabIndex        =   2
            Top             =   300
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Informe de Clasificación de Deudores y Provisiones"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   570
            TabIndex        =   3
            Top             =   30
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Anexo N° 5"
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
            Picture         =   "GesCtb_frm_815.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   780
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
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
            Left            =   4770
            Picture         =   "GesCtb_frm_815.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_815.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_815.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   825
         Left            =   30
         TabIndex        =   8
         Top             =   1470
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   60
            Width           =   3795
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1200
            TabIndex        =   10
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
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   450
            Width           =   1065
         End
         Begin VB.Label Label5 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   90
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_str_FecIni        As String
Dim r_str_FecFin        As String

Private Sub cmd_ExpExc_Click()
      
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
   
   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01"
   r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
   
   Call fs_GenExc(r_str_FecIni, r_str_FecFin)
   
End Sub


Private Sub fs_GenExc(ByVal p_FecIni As String, ByVal p_FecFin As String)
   
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   
   Dim i As Integer
   Dim j As Integer
   
   Dim r_arr_DetGar()      As modprc_g_tpo_DetGar
      
   Dim r_dbl_PvCoMv_0         As Double
   Dim r_dbl_PvCoBi_0         As Double
   Dim r_dbl_PvCoFi_0         As Double
   Dim r_dbl_SalMiv_0         As Double
   Dim r_dbl_SalBid_0         As Double
   Dim r_dbl_SalFid_0         As Double
   
   Dim r_dbl_PvCoMv_1         As Double
   Dim r_dbl_PvCoBi_1         As Double
   Dim r_dbl_PvCoFi_1         As Double
   Dim r_dbl_SalMiv_1         As Double
   Dim r_dbl_SalBid_1         As Double
   Dim r_dbl_SalFid_1         As Double
      
   Dim r_dbl_PvCoMv_2         As Double
   Dim r_dbl_PvCoBi_2         As Double
   Dim r_dbl_PvCoFi_2         As Double
   Dim r_dbl_SalMiv_2         As Double
   Dim r_dbl_SalBid_2         As Double
   Dim r_dbl_SalFid_2         As Double
   
   Dim r_dbl_PvCoMv_3         As Double
   Dim r_dbl_PvCoBi_3         As Double
   Dim r_dbl_PvCoFi_3         As Double
   Dim r_dbl_SalMiv_3         As Double
   Dim r_dbl_SalBid_3         As Double
   Dim r_dbl_SalFid_3         As Double
   
   Dim r_dbl_PvCoMv_4         As Double
   Dim r_dbl_PvCoBi_4         As Double
   Dim r_dbl_PvCoFi_4         As Double
   Dim r_dbl_SalMiv_4         As Double
   Dim r_dbl_SalBid_4         As Double
   Dim r_dbl_SalFid_4         As Double
   
   Dim r_dbl_PvRqMv_0         As Double
   
   Dim r_int_ClaGar        As Integer
   Dim r_int_Contad        As Integer
   
   Dim c As Integer
   
   'Leer Tabla de Creditos del mes CRE_HIPCIE
   g_str_Parame = "SELECT HIPCIE_NUMOPE, HIPCIE_CLAPRV, HIPCIE_PRVGEN, HIPCIE_PRVESP, HIPCIE_PRVCIC, HIPCIE_SITCRE, HIPCIE_CLACRE, HIPCIE_SALCAP, HIPCIE_SALCON, HIPCIE_TIPCAM, HIPMAE_GARLIN, HIPMAE_CODPRD, HIPCIE_TIPGAR FROM CRE_HIPCIE H, CRE_HIPMAE M WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_NUMOPE = HIPMAE_NUMOPE AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_NUMOPE ASC "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Saldos para generar el Reporte.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      
      .Cells(5, 2) = "ANEXO Nº 5-D"
      .Cells(6, 2) = "INFORME DE CLASIFICACION DE DEUDORES Y PROVISIONES"
      .Cells(8, 2) = "EDPYME MICASITA"
      .Cells(8, 8) = "CODIGO: 00240"
      .Cells(10, 2) = "Al " & Right(p_FecFin, 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & Left(p_FecFin, 4)
      .Cells(11, 2) = "(En nuevos soles)"
      .Cells(13, 2) = "INFORME DE CLASIFICACIÓN DE LOS DEUDORES DE LA CARTERA DE CRÉDITOS, CONTINGENTES Y ARRENDAMIENTOS"
      .Cells(14, 2) = "FINANCIEROS QUE RESPALDAN FINANCIAMIENTOS O LINEAS DE CREDITO 1/"
      
      .Cells(107, 3) = "GERENTE"
      .Cells(107, 5) = "CONTADOR"
      .Cells(107, 7) = "FUNCIONARIO"
           
      .Cells(108, 3) = "GENERAL"
      .Cells(108, 5) = "GENERAL"
      .Cells(108, 7) = "RESPONSABLE"
      
      .Cells(107, 3).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(107, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(108, 3).HorizontalAlignment = xlHAlignCenter
      
      .Cells(107, 5).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(107, 5).HorizontalAlignment = xlHAlignCenter
      .Cells(108, 5).HorizontalAlignment = xlHAlignCenter
      
      .Cells(107, 7).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(107, 7).HorizontalAlignment = xlHAlignCenter
      .Cells(108, 7).HorizontalAlignment = xlHAlignCenter
      
      .Cells(16, 2) = "Línea Mivivienda:"
      .Cells(16, 2).Font.Bold = True
      
      .Cells(45, 2) = "Línea BID:"
      .Cells(45, 2).Font.Bold = True
      
      .Cells(74, 2) = "Línea Fidel Ramirez Prado:"
      .Cells(74, 2).Font.Bold = True
                       
      'LINEA MIVIVIENDA
      For i = 18 To 101
                  
         If i = 38 Then
            .Cells(i, 2) = "TOTAL"
            i = i + 1
         ElseIf i = 67 Then
            .Cells(i, 2) = "TOTAL"
            i = i + 1
         ElseIf i = 96 Then
            .Cells(i, 2) = "TOTAL"
            i = i + 1
         ElseIf i = 44 Or i = 73 Then
            i = i + 3
         End If
                  
         If i = 18 Or i = 47 Or i = 76 Then
             
            .Cells(i, 2) = "A.- MONTO DE LOS CRÉDITOS, CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS 2/"
            .Range(.Cells(i, 3), .Cells(i, 8)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(i, 2), .Cells(i, 8)).VerticalAlignment = xlCenter
            .Range(.Cells(i, 2), .Cells(i, 8)).Font.Bold = True
            .Range(.Cells(i, 2), .Cells(i, 8)).Borders.Color = RGB(0, 0, 0)
             
         ElseIf i = 23 Or i = 52 Or i = 81 Then
            
            .Cells(i, 2) = "B.- PROVISIONES CONSTITUIDAS 3/"
            .Range(.Cells(i, 3), .Cells(i, 8)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(i, 2), .Cells(i, 8)).VerticalAlignment = xlCenter
            .Range(.Cells(i, 2), .Cells(i, 8)).Font.Bold = True
            .Range(.Cells(i, 2), .Cells(i, 8)).Borders.Color = RGB(0, 0, 0)
         
         ElseIf i = 28 Or i = 57 Or i = 86 Then
            
            .Cells(i, 2) = "C.- PROVISIONES REQUERIDAS 4/"
            .Range(.Cells(i, 3), .Cells(i, 8)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(i, 2), .Cells(i, 8)).VerticalAlignment = xlCenter
            .Range(.Cells(i, 2), .Cells(i, 8)).Font.Bold = True
            .Range(.Cells(i, 2), .Cells(i, 8)).Borders.Color = RGB(0, 0, 0)
         
         ElseIf i = 33 Or i = 62 Or i = 91 Then
            
            .Cells(i, 2) = "D.- SUPERÁVIT (DÉFICIT) DE PROVISIONES 5/"
            .Range(.Cells(i, 3), .Cells(i, 8)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(i, 2), .Cells(i, 8)).VerticalAlignment = xlCenter
            .Range(.Cells(i, 2), .Cells(i, 8)).Font.Bold = True
            .Range(.Cells(i, 2), .Cells(i, 8)).Borders.Color = RGB(0, 0, 0)
            
            .Range(.Cells(i + 5, 2), .Cells(i + 5, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(i + 5, 2), .Cells(i + 5, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(i + 5, 2), .Cells(i + 5, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(i + 5, 2), .Cells(i + 5, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous

         
         ElseIf i = 39 Or i = 68 Or i = 97 Then
            
            .Cells(i, 2) = "E.- MONTO DE LOS CREDITOS, CONTINGENTES 6/"
            .Range(.Cells(i, 3), .Cells(i, 8)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(i, 2), .Cells(i, 8)).VerticalAlignment = xlCenter
            .Range(.Cells(i, 2), .Cells(i, 8)).Font.Bold = True
            .Range(.Cells(i, 2), .Cells(i, 8)).Borders.Color = RGB(0, 0, 0)
            
            .Range(.Cells(i + 4, 2), .Cells(i + 4, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         
         End If
         
         .Cells(i, 3) = "Normal"
         .Cells(i, 4) = "CPP"
         .Cells(i, 5) = "Deficiente"
         .Cells(i, 6) = "Dudoso"
         .Cells(i, 7) = "Pérdida"
         .Cells(i, 8) = "Total"
                 
         i = i + 4
                 
      Next i
              
      For j = 19 To 101
         
         If j = 39 Or j = 68 Or j = 97 Then
            j = j + 1
         End If
         
         .Cells(j + 0, 2) = "Comerciales"
         .Cells(j + 1, 2) = "MES"
         .Cells(j + 2, 2) = "Hipotecario para Vivienda"
         .Cells(j + 3, 2) = "Consumo"
         
         For c = 0 To 3
            
            .Range(.Cells(j + c, 2), .Cells(j + c, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(j + c, 2), .Cells(j + c, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(j + c, 2), .Cells(j + c, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
         
         Next c
                   
         If j = 40 Or j = 69 Then
            j = j + 7
         Else
            j = j + 4
         End If
                 
      Next j
        
      .Range(.Range("B2:H2"), .Range("B14:H14")).Font.Bold = True
      .Range(.Range("B2:H2"), .Range("B14:H14")).HorizontalAlignment = xlHAlignCenter
            
      .Range("B5:H5").Merge
      .Range("B6:H6").Merge
      .Range("B10:H10").Merge
      .Range("B11:H11").Merge
      .Range("B13:H13").Merge
      .Range("B14:H14").Merge
      
      .Range("B4").HorizontalAlignment = xlHAlignLeft
           
      .Columns("A").ColumnWidth = 2
      .Columns("B").ColumnWidth = 76
      .Columns("C").ColumnWidth = 15
      .Columns("D").ColumnWidth = 15
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 15
      .Columns("G").ColumnWidth = 15
      .Columns("H").ColumnWidth = 15
      
      .Range(.Cells(1, 1), .Cells(400, 400)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(400, 400)).Font.Size = 10
      
                 
   End With
   
   g_rst_Princi.MoveFirst
     
   Do While Not g_rst_Princi.EOF
      
      If g_rst_Princi!HIPMAE_GARLIN = "000001" Then
         
         If g_rst_Princi!HIPMAE_CODPRD = "003" Or g_rst_Princi!HIPMAE_CODPRD = "004" Or g_rst_Princi!HIPMAE_CODPRD = "007" Then
            'Vigente
            If g_rst_Princi!HIPCIE_SITCRE = 1 Then
               'Normal
               If g_rst_Princi!HIPCIE_CLACRE = 0 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoMv_0 = r_dbl_PvCoMv_0 + (g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVCIC)
                  r_dbl_SalMiv_0 = r_dbl_SalMiv_0 + g_rst_Princi!HIPCIE_SALCAP
                  
                  'CALCULO DE PROVISIONES REQUERIDAS
                  
               End If
            'Vencido
            ElseIf g_rst_Princi!HIPCIE_SITCRE = 5 Then
               'CPP
               If g_rst_Princi!HIPCIE_CLACRE = 1 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoMv_1 = r_dbl_PvCoMv_1 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalMiv_1 = r_dbl_SalMiv_1 + g_rst_Princi!HIPCIE_SALCAP
                  
                  'CALCULO DE PROVISIONES REQUERIDAS
                  
               'Deficiente
               ElseIf g_rst_Princi!HIPCIE_CLACRE = 2 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoMv_2 = r_dbl_PvCoMv_2 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalMiv_2 = r_dbl_SalMiv_2 + g_rst_Princi!HIPCIE_SALCAP
               'Dudoso
               ElseIf g_rst_Princi!HIPCIE_CLACRE = 3 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoMv_3 = r_dbl_PvCoMv_3 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalMiv_3 = r_dbl_SalMiv_3 + g_rst_Princi!HIPCIE_SALCAP
               'Perdida
               ElseIf g_rst_Princi!HIPCIE_CLACRE = 4 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoMv_4 = r_dbl_PvCoMv_4 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalMiv_4 = r_dbl_SalMiv_4 + g_rst_Princi!HIPCIE_SALCAP
               End If
            End If
         End If
      ElseIf g_rst_Princi!HIPMAE_GARLIN = "000002" Then
         
         If g_rst_Princi!HIPMAE_CODPRD = "001" Or g_rst_Princi!HIPMAE_CODPRD = "002" Then
            If g_rst_Princi!HIPCIE_SITCRE = 1 Then
               If g_rst_Princi!HIPCIE_CLACRE = 0 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoBi_0 = r_dbl_PvCoBi_0 + (g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVCIC)
                  r_dbl_SalBid_0 = r_dbl_SalBid_0 + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               
                  'CALCULO DE PROVISIONES REQUERIDAS
               
               End If
            ElseIf g_rst_Princi!HIPCIE_SITCRE = 5 Then
               If g_rst_Princi!HIPCIE_CLACRE = 1 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoBi_1 = r_dbl_PvCoBi_1 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalBid_1 = r_dbl_SalBid_1 + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  
                  'CALCULO DE PROVISIONES REQUERIDAS
               
               ElseIf g_rst_Princi!HIPCIE_CLACRE = 2 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoBi_2 = r_dbl_PvCoBi_2 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalBid_2 = r_dbl_SalBid_2 + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               
                  'CALCULO DE PROVISIONES REQUERIDAS
                  
               ElseIf g_rst_Princi!HIPCIE_CLACRE = 3 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoBi_3 = r_dbl_PvCoBi_3 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalBid_3 = r_dbl_SalBid_3 + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               
                  'CALCULO DE PROVISIONES REQUERIDAS
               
               ElseIf g_rst_Princi!HIPCIE_CLACRE = 4 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoBi_4 = r_dbl_PvCoBi_4 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalBid_4 = r_dbl_SalBid_4 + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               
                  'CALCULO DE PROVISIONES REQUERIDAS
               
               End If
            End If
         End If
      
      ElseIf g_rst_Princi!HIPMAE_GARLIN = "000003" Then
         
         If g_rst_Princi!HIPMAE_CODPRD = "002" Then
            If g_rst_Princi!HIPCIE_SITCRE = 1 Then
               If g_rst_Princi!HIPCIE_CLACRE = 0 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoFi_0 = r_dbl_PvCoFi_0 + (g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVCIC)
                  r_dbl_SalFid_0 = r_dbl_SalFid_0 + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               End If
            ElseIf g_rst_Princi!HIPCIE_SITCRE = 5 Then
               If g_rst_Princi!HIPCIE_CLACRE = 1 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoFi_1 = r_dbl_PvCoFi_1 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalFid_1 = r_dbl_SalFid_1 + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               ElseIf g_rst_Princi!HIPCIE_CLACRE = 2 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoFi_2 = r_dbl_PvCoFi_2 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalFid_2 = r_dbl_SalFid_2 + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               ElseIf g_rst_Princi!HIPCIE_CLACRE = 3 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoFi_3 = r_dbl_PvCoFi_3 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalFid_3 = r_dbl_SalFid_3 + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               ElseIf g_rst_Princi!HIPCIE_CLACRE = 4 Then
                  'CALCULO DE PROVISIONES CONSTITUIDAS
                  r_dbl_PvCoFi_4 = r_dbl_PvCoFi_4 + g_rst_Princi!HIPCIE_PRVESP
                  r_dbl_SalFid_4 = r_dbl_SalFid_4 + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               End If
            End If
         End If
      End If
       
      'r_int_ConVer = r_int_ConVer + 1
        
      g_rst_Princi.MoveNext
      DoEvents
   
   Loop
   
   'Mi Vivienda
   r_obj_Excel.ActiveSheet.Cells(21, 3) = Format(r_dbl_SalMiv_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(21, 4) = Format(r_dbl_SalMiv_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(21, 5) = Format(r_dbl_SalMiv_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(21, 6) = Format(r_dbl_SalMiv_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(21, 7) = Format(r_dbl_SalMiv_4, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(21, 8) = Format(r_dbl_SalMiv_0 + r_dbl_SalMiv_1 + r_dbl_SalMiv_2 + r_dbl_SalMiv_3 + r_dbl_SalMiv_4, "###,###,##0.00")
  
   r_obj_Excel.ActiveSheet.Cells(26, 3) = Format(r_dbl_PvCoMv_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(26, 4) = Format(r_dbl_PvCoMv_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(26, 5) = Format(r_dbl_PvCoMv_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(26, 6) = Format(r_dbl_PvCoMv_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(26, 7) = Format(r_dbl_PvCoMv_4, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(26, 8) = Format(r_dbl_PvCoMv_0 + r_dbl_PvCoMv_1 + r_dbl_PvCoMv_2 + r_dbl_PvCoMv_3 + r_dbl_PvCoMv_4, "###,###,##0.00")
   
   r_obj_Excel.ActiveSheet.Cells(31, 3) = Format(r_dbl_PvCoMv_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(31, 4) = Format(r_dbl_PvCoMv_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(31, 5) = Format(r_dbl_PvCoMv_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(31, 6) = Format(r_dbl_PvCoMv_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(31, 7) = Format(r_dbl_PvCoMv_4, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(31, 8) = Format(r_dbl_PvCoMv_0 + r_dbl_PvCoMv_1 + r_dbl_PvCoMv_2 + r_dbl_PvCoMv_3 + r_dbl_PvCoMv_4, "###,###,##0.00")
   
   r_obj_Excel.ActiveSheet.Cells(36, 3) = Format(r_dbl_PvCoMv_0 - r_dbl_PvCoMv_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(36, 4) = Format(r_dbl_PvCoMv_1 - r_dbl_PvCoMv_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(36, 5) = Format(r_dbl_PvCoMv_2 - r_dbl_PvCoMv_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(36, 6) = Format(r_dbl_PvCoMv_3 - r_dbl_PvCoMv_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(36, 7) = Format(r_dbl_PvCoMv_4 - r_dbl_PvCoMv_4, "###,###,##0.00")
         
   'BID
   r_obj_Excel.ActiveSheet.Cells(50, 3) = Format(r_dbl_SalBid_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(50, 4) = Format(r_dbl_SalBid_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(50, 5) = Format(r_dbl_SalBid_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(50, 6) = Format(r_dbl_SalBid_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(50, 7) = Format(r_dbl_SalBid_4, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(50, 8) = Format(r_dbl_SalBid_0 + r_dbl_SalBid_1 + r_dbl_SalBid_2 + r_dbl_SalBid_3 + r_dbl_SalBid_4, "###,###,##0.00")
   
   r_obj_Excel.ActiveSheet.Cells(55, 3) = Format(r_dbl_PvCoBi_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(55, 4) = Format(r_dbl_PvCoBi_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(55, 5) = Format(r_dbl_PvCoBi_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(55, 6) = Format(r_dbl_PvCoBi_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(55, 7) = Format(r_dbl_PvCoBi_4, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(55, 8) = Format(r_dbl_PvCoBi_0 + r_dbl_PvCoBi_1 + r_dbl_PvCoBi_2 + r_dbl_PvCoBi_3 + r_dbl_PvCoBi_4, "###,###,##0.00")
     
   r_obj_Excel.ActiveSheet.Cells(60, 3) = Format(r_dbl_PvCoBi_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(60, 4) = Format(r_dbl_PvCoBi_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(60, 5) = Format(r_dbl_PvCoBi_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(60, 6) = Format(r_dbl_PvCoBi_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(60, 7) = Format(r_dbl_PvCoBi_4, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(60, 8) = Format(r_dbl_PvCoBi_0 + r_dbl_PvCoBi_1 + r_dbl_PvCoBi_2 + r_dbl_PvCoBi_3 + r_dbl_PvCoBi_4, "###,###,##0.00")
     
   r_obj_Excel.ActiveSheet.Cells(65, 3) = Format(r_dbl_PvCoBi_0 - r_dbl_PvCoBi_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(65, 4) = Format(r_dbl_PvCoBi_1 - r_dbl_PvCoBi_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(65, 5) = Format(r_dbl_PvCoBi_2 - r_dbl_PvCoBi_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(65, 6) = Format(r_dbl_PvCoBi_3 - r_dbl_PvCoBi_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(65, 7) = Format(r_dbl_PvCoBi_4 - r_dbl_PvCoBi_4, "###,###,##0.00")
     
   'Fidel Ramirez
   r_obj_Excel.ActiveSheet.Cells(79, 3) = Format(r_dbl_SalFid_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(79, 4) = Format(r_dbl_SalFid_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(79, 5) = Format(r_dbl_SalFid_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(79, 6) = Format(r_dbl_SalFid_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(79, 7) = Format(r_dbl_SalFid_4, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(79, 8) = Format(r_dbl_SalFid_0 + r_dbl_SalFid_1 + r_dbl_SalFid_2 + r_dbl_SalFid_3 + r_dbl_SalFid_4, "###,###,##0.00")
    
   r_obj_Excel.ActiveSheet.Cells(84, 3) = Format(r_dbl_PvCoFi_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(84, 4) = Format(r_dbl_PvCoFi_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(84, 5) = Format(r_dbl_PvCoFi_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(84, 6) = Format(r_dbl_PvCoFi_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(84, 7) = Format(r_dbl_PvCoFi_4, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(84, 8) = Format(r_dbl_PvCoFi_0 + r_dbl_PvCoFi_1 + r_dbl_PvCoFi_2 + r_dbl_PvCoFi_3 + r_dbl_PvCoFi_4, "###,###,##0.00")
     
   r_obj_Excel.ActiveSheet.Cells(89, 3) = Format(r_dbl_PvCoFi_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(89, 4) = Format(r_dbl_PvCoFi_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(89, 5) = Format(r_dbl_PvCoFi_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(89, 6) = Format(r_dbl_PvCoFi_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(89, 7) = Format(r_dbl_PvCoFi_4, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(89, 8) = Format(r_dbl_PvCoFi_0 + r_dbl_PvCoFi_1 + r_dbl_PvCoFi_2 + r_dbl_PvCoFi_3 + r_dbl_PvCoFi_4, "###,###,##0.00")
     
   r_obj_Excel.ActiveSheet.Cells(94, 3) = Format(r_dbl_PvCoFi_0 - r_dbl_PvCoFi_0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(94, 4) = Format(r_dbl_PvCoFi_1 - r_dbl_PvCoFi_1, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(94, 5) = Format(r_dbl_PvCoFi_2 - r_dbl_PvCoFi_2, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(94, 6) = Format(r_dbl_PvCoFi_3 - r_dbl_PvCoFi_3, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(94, 7) = Format(r_dbl_PvCoFi_4 - r_dbl_PvCoFi_4, "###,###,##0.00")
   
   r_obj_Excel.ActiveSheet.Cells(38, 3) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(38, 4) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(38, 5) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(38, 6) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(38, 7) = Format(0, "###,###,##0.00")
   
   r_obj_Excel.ActiveSheet.Cells(67, 3) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(67, 4) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(67, 5) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(67, 6) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(67, 7) = Format(0, "###,###,##0.00")
   
   r_obj_Excel.ActiveSheet.Cells(96, 3) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(96, 4) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(96, 5) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(96, 6) = Format(0, "###,###,##0.00")
   r_obj_Excel.ActiveSheet.Cells(96, 7) = Format(0, "###,###,##0.00")
     
     
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing

End Sub


