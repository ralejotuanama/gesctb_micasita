VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptCtb_13 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   6720
   ClientTop       =   6675
   ClientWidth     =   6075
   Icon            =   "GesCtb_frm_846.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2775
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6075
      _Version        =   65536
      _ExtentX        =   10716
      _ExtentY        =   4895
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
         TabIndex        =   6
         Top             =   60
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   270
            Left            =   630
            TabIndex        =   7
            Top             =   120
            Width           =   5205
            _Version        =   65536
            _ExtentX        =   9181
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Provisión de clientes Morosos y Alineados"
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
            Picture         =   "GesCtb_frm_846.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   8
         Top             =   780
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
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
            Left            =   30
            Picture         =   "GesCtb_frm_846.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5340
            Picture         =   "GesCtb_frm_846.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1245
         Left            =   60
         TabIndex        =   9
         Top             =   1470
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   2196
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
         Begin VB.CheckBox Chk_FecAct 
            Caption         =   "A la Fecha"
            Height          =   285
            Left            =   1170
            TabIndex        =   2
            Top             =   870
            Width           =   1995
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   135
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1170
            TabIndex        =   1
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
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   195
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   520
            Width           =   885
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_PerMes        As String
Dim l_str_PerAno        As String
Dim l_str_MesAnt        As String
Dim l_str_AnoAnt        As String
Dim l_str_MesRcc        As String
Dim l_str_AnoRcc        As String

Private Sub Chk_FecAct_Click()
   If Chk_FecAct.Value = 1 Then
      cmb_PerMes.ListIndex = -1
      cmb_PerMes.Enabled = False
      ipp_PerAno.Value = Year(date)
      ipp_PerAno.Enabled = False
      l_str_PerMes = Month(date)
      l_str_PerAno = Year(date)
      Call gs_SetFocus(cmd_ExpExc)
   ElseIf Chk_FecAct.Value = 0 Then
      cmb_PerMes.Enabled = True
      ipp_PerAno.Enabled = True
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If Chk_FecAct.Value = 1 Then
       'Confirmacion
       If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Exit Sub
       End If
       
       Screen.MousePointer = 11
       Me.Enabled = False
       Call fs_GenExc_Actual_Nuevo
       Me.Enabled = True
       Screen.MousePointer = 0
        
   ElseIf Chk_FecAct.Value = 0 Then
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
       
       'Confirmacion
       If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Exit Sub
       End If
       
       Screen.MousePointer = 11
       Me.Enabled = False
       
       l_str_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
       l_str_PerAno = ipp_PerAno.Text
       l_str_MesAnt = IIf(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 1, 12, cmb_PerMes.ItemData(cmb_PerMes.ListIndex) - 1)
       l_str_AnoAnt = IIf(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 1, CLng(ipp_PerAno.Text) - 1, ipp_PerAno.Text)
       Call fs_Recorset_nc
       Call fs_GenExc(l_str_PerMes, l_str_PerAno, l_str_MesAnt, l_str_AnoAnt)
       
       Me.Enabled = True
       Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Caption = modgen_g_str_NomPlt
    
    Call gs_CentraForm(Me)
    Call fs_Inicia
     
    Call gs_SetFocus(cmb_PerMes)
    Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
    Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
    ipp_PerAno.Text = Year(date)
End Sub

Private Sub fs_Recorset_nc()
    Set g_rst_Listas = New ADODB.Recordset
    g_rst_Listas.Fields.Append "MES", adBigInt, 3, adFldFixed
    g_rst_Listas.Fields.Append "ANNO", adBigInt, 4, adFldIsNullable
    g_rst_Listas.Fields.Append "TIPO", adChar, 100, adFldIsNullable
    g_rst_Listas.Fields.Append "CANTIDAD", adBigInt, 3, adFldFixed
    g_rst_Listas.Open , , adOpenKeyset, adLockOptimistic
End Sub

Private Sub fs_GenExc(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer)
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_Nindex     As Integer
Dim r_dbl_SumCapBal  As Double
Dim r_dbl_SumCapAju  As Double
Dim r_dbl_SumPrv     As Double
Dim r_dbl_SumVol     As Double
Dim r_dbl_SumNor     As Double
Dim r_dbl_TotCapBal  As Double
Dim r_dbl_TotCapAju  As Double
Dim r_dbl_TotPrv     As Double
Dim r_dbl_TotVol     As Double
Dim r_int_nroaux     As Integer
Dim r_bol_FlagOp     As Boolean
Dim r_dbl_NvDHip     As Double
Dim r_dbl_NvDCom     As Double
Dim r_dbl_PorHip1    As Double
Dim r_dbl_PorHip2    As Double
Dim r_dbl_xAmCom     As Double
Dim r_dbl_CMoros     As Double
Dim r_dbl_HipPrv     As Double
Dim r_dbl_CliMoI     As Double
Dim r_dbl_CMoDia     As Double
Dim r_dbl_PrvNet     As Double
Dim l_str_RccMes     As String
Dim l_str_RccAno     As String
Dim l_int_NumEnt     As Integer
Dim l_dbl_MtoDeu     As Double
Dim r_str_Cadena     As String
Dim r_bol_Flag1      As Boolean
Dim r_bol_Flag2      As Boolean
Dim r_int_NumMes     As Integer
Dim p_MesActN        As Integer
Dim p_AnoActN        As Integer
Dim p_MesAntN        As Integer
Dim p_AnoAntN        As Integer
Dim p_str_CadAux     As String
Dim r_rst_GenAux     As ADODB.Recordset

   r_dbl_SumCapAju = 0
   r_dbl_SumPrv = 0
   r_dbl_SumVol = 0
   r_dbl_SumNor = 0
   r_dbl_TotCapBal = 0
   r_dbl_TotCapAju = 0
   r_dbl_TotPrv = 0
   r_dbl_TotVol = 0
   r_int_Nindex = 0
   
   'inicializar valores
   r_dbl_NvDHip = 0
   r_dbl_NvDCom = 0
   r_dbl_PorHip1 = 0
   r_dbl_PorHip2 = 0
   r_dbl_xAmCom = 0
   r_dbl_CMoros = 0
   r_dbl_HipPrv = 0
   r_dbl_CliMoI = 0
   r_dbl_CMoDia = 0
   r_dbl_PrvNet = 0
   
   r_bol_FlagOp = False
   g_str_Parame = gf_Query(p_MesAct, p_AnoAct, p_MesAnt, p_AnoAnt)
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "MOROSOS"
   r_int_NroFil = 2
   
   With r_obj_Excel.Sheets(1)
      .Range("A" & r_int_NroFil & ":W" & r_int_NroFil & "").Merge
      .Range("A" & r_int_NroFil & ":W" & r_int_NroFil & "").Font.Bold = True
      .Range("A" & r_int_NroFil & ":W" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_NroFil & ":W" & r_int_NroFil & "").Font.Size = 14
      .Range("A" & r_int_NroFil & ":W" & r_int_NroFil & "").Font.Underline = xlUnderlineStyleSingle
      
      .Cells(r_int_NroFil, 1) = "PROVISIÓN DE CLIENTES MOROSOS Y ALINEADOS AL MES DE " & Trim(cmb_PerMes.Text) & " DEL " & Trim(ipp_PerAno.Text)
      r_int_NroFil = r_int_NroFil + 2
      
      .Range("M" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Range("Q" & r_int_NroFil & ":T" & r_int_NroFil & "").Merge
      .Range("M" & r_int_NroFil & ":P" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("Q" & r_int_NroFil & ":T" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("M" & r_int_NroFil & ":P" & r_int_NroFil & "").Font.Bold = True
      .Range("Q" & r_int_NroFil & ":T" & r_int_NroFil & "").Font.Bold = True
      
      .Cells(r_int_NroFil, 13) = "ACTUAL"
      .Cells(r_int_NroFil, 17) = "ANTERIOR"
      
      r_int_NroFil = r_int_NroFil + 1
      .Cells(r_int_NroFil, 1) = "PRODUCTO"
      .Cells(r_int_NroFil, 2) = "Nº OPERACION"
      .Cells(r_int_NroFil, 3) = "TIPO"
      .Cells(r_int_NroFil, 4) = "NOMBRE DEL CLIENTE"
      .Cells(r_int_NroFil, 5) = "DÍAS ATRASO"
      .Cells(r_int_NroFil, 6) = "CUOTAS ATR"
      .Cells(r_int_NroFil, 7) = "D36 o P24"
      .Cells(r_int_NroFil, 8) = "GARANTÍA"
      .Cells(r_int_NroFil, 9) = "VAL. GARANTÍA"
      .Cells(r_int_NroFil, 10) = "CAPITAL BAL"
      .Cells(r_int_NroFil, 11) = "CAPITAL AJU"
      .Cells(r_int_NroFil, 12) = "TASA %"
      .Cells(r_int_NroFil, 13) = "PROVISIÓN"
      .Cells(r_int_NroFil, 14) = "VOLUNTARIA"
      .Cells(r_int_NroFil, 15) = "CLA. ALINEADA"
      .Cells(r_int_NroFil, 16) = "CLA. INTERNA"
      .Cells(r_int_NroFil, 17) = "PROVISIÓN"
      .Cells(r_int_NroFil, 18) = "VOLUNTARIA"
      .Cells(r_int_NroFil, 19) = "CLA. ALINEADA"
      .Cells(r_int_NroFil, 20) = "CLA. INTERNA"
      .Cells(r_int_NroFil, 21) = "AJUSTE"
      .Cells(r_int_NroFil, 22) = "COB. FMV"
      .Cells(r_int_NroFil, 23) = "COB. FMV RC"
      .Cells(r_int_NroFil, 24) = "CUO.PAG. > 30 DIAS"
      .Cells(r_int_NroFil, 25) = "N° ENT. REPORT."
      .Cells(r_int_NroFil, 26) = "TOTAL DEUDA"
      .Cells(r_int_NroFil, 27) = "CODIGO EXCEPCION"
      .Cells(r_int_NroFil, 28) = "NOMBRE DEL PROYECTO"
      .Cells(r_int_NroFil, 29) = "MICROEMPRESARIO"
      .Cells(r_int_NroFil, 30) = "HIPOTECA MATRIZ"
      .Cells(r_int_NroFil, 31) = "FECHA COMPROMISO"
      .Cells(r_int_NroFil, 32) = "TIPO COMPROMISO"
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 32)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 32)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 11
      .Columns("B").ColumnWidth = 14
      .Columns("C").ColumnWidth = 12
      .Columns("D").ColumnWidth = 46
      .Columns("E").ColumnWidth = 12
      .Columns("F").ColumnWidth = 12
      .Columns("G").ColumnWidth = 12
      .Columns("H").ColumnWidth = 12
      .Columns("I").ColumnWidth = 15
      .Columns("J").ColumnWidth = 12
      .Columns("K").ColumnWidth = 12
      .Columns("L").ColumnWidth = 8
      .Columns("M").ColumnWidth = 12
      .Columns("N").ColumnWidth = 12
      .Columns("O").ColumnWidth = 14
      .Columns("P").ColumnWidth = 13
      .Columns("Q").ColumnWidth = 12
      .Columns("R").ColumnWidth = 12
      .Columns("S").ColumnWidth = 14
      .Columns("T").ColumnWidth = 13
      .Columns("U").ColumnWidth = 11
      .Columns("V").ColumnWidth = 11
      .Columns("W").ColumnWidth = 13
      .Columns("X").ColumnWidth = 18
      .Columns("Y").ColumnWidth = 15
      .Columns("Z").ColumnWidth = 14
      .Columns("AA").ColumnWidth = 18
      .Columns("AB").ColumnWidth = 40
      .Columns("AC").ColumnWidth = 19
      .Columns("AD").ColumnWidth = 19
      .Columns("AE").ColumnWidth = 20
      .Columns("AF").ColumnWidth = 33
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      .Columns("T").HorizontalAlignment = xlHAlignCenter
      .Columns("X").HorizontalAlignment = xlHAlignCenter
      .Columns("Y").HorizontalAlignment = xlHAlignCenter
      .Columns("AA").HorizontalAlignment = xlHAlignCenter
      .Columns("AC").HorizontalAlignment = xlHAlignCenter
      .Columns("AD").HorizontalAlignment = xlHAlignCenter
      .Columns("AE").HorizontalAlignment = xlHAlignCenter
      .Columns("AF").HorizontalAlignment = xlHAlignLeft
      
      Do While r_int_Nindex <= 4
         r_int_nroaux = 1
         
         Do While r_int_nroaux < 3
            r_dbl_SumCapBal = 0
            r_dbl_SumCapAju = 0
            r_dbl_SumPrv = 0
            r_dbl_SumVol = 0
            g_rst_Princi.Filter = adFilterNone
            g_rst_Princi.MoveFirst
            r_int_NroFil = r_int_NroFil + 1
            
            If r_int_nroaux = 1 Then
               g_rst_Princi.Filter = "HIPCIE_CLAPRV = " & r_int_Nindex & " AND HIPCIE_TIPGAR > 2"
            Else
               g_rst_Princi.Filter = "HIPCIE_CLAPRV = " & r_int_Nindex & " AND HIPCIE_TIPGAR < 3"
            End If
            
            If g_rst_Princi.EOF Then
               r_bol_FlagOp = True
            End If
            p_str_CadAux = ""
            Do While Not g_rst_Princi.EOF
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 1) = Trim(g_rst_Princi!PRODUCTO)
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 2) = gf_Formato_NumOpe(g_rst_Princi!OPERACION)
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 3) = fs_MuestraTipo(g_rst_Princi!CLASIFICACION, g_rst_Princi!CLASIFICACION2)
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 4) = Trim(g_rst_Princi!NOMBRECLIENTE)
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5) = g_rst_Princi!DIASATRASO
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5)).NumberFormat = "#,##0"
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 6) = fs_CuotasAtrasadas_Cierre(g_rst_Princi!OPERACION, p_AnoAct, p_MesAct)
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 6), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 6)).NumberFormat = "##0"
              
               p_str_CadAux = fs_Clasifica_Dudoso_Perdida(g_rst_Princi!OPERACION, p_AnoAct, p_MesAct)
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 7) = p_str_CadAux
               
               If InStr(p_str_CadAux, "D") > 0 Then
                  If CInt(Mid(CStr(p_str_CadAux), 1, InStr(p_str_CadAux, "D") - 1)) > 36 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 7).Font.Color = -16776961
                     r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 7).Font.Bold = True
                  End If
               ElseIf InStr(p_str_CadAux, "P") > 0 Then
                  If CInt(Mid(CStr(p_str_CadAux), 1, InStr(p_str_CadAux, "P") - 1)) > 24 Then
                     r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 7).Font.Color = -16776961
                     r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 7).Font.Bold = True
                  End If
               End If
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 8) = Trim(g_rst_Princi!TIPOGARAN)
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 9) = CDbl(Format(g_rst_Princi!VALORGARANTIA, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 9), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 9)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10) = CDbl(Format(g_rst_Princi!CAPITAL_BAL, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11) = CDbl(Format(g_rst_Princi!CAPITAL, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 12) = CDbl(Format(fs_Obtiene_TasaProvision(g_rst_Princi!OPERACION, g_rst_Princi!HIPCIE_CLACLI, g_rst_Princi!HIPCIE_TIPGAR, g_rst_Princi!TASA), "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 12), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 12)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13) = CDbl(Format(g_rst_Princi!PROVISION, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14) = Trim(g_rst_Princi!PROV_VOLUNT1)
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 15) = Trim(g_rst_Princi!CLASIFICACION)
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 16) = Trim(g_rst_Princi!CLAINT1)
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 17) = CDbl(Format(g_rst_Princi!PROVISION2, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 17), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 17)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 18) = Trim(g_rst_Princi!PROV_VOLUNT2)
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 18), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 18)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 19) = Trim(g_rst_Princi!CLASIFICACION2)
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 20) = Trim(g_rst_Princi!CLAINT2)
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 21) = CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 21), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 21)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 22) = CDbl(Format(g_rst_Princi!HIPCIE_CBRFMV, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 22), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 22)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 23) = CDbl(Format(g_rst_Princi!HIPCIE_CBRFMV_RC, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 23), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 23)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 24) = "'" & fs_Calcula_CuotasPagadasConAtraso(g_rst_Princi!OPERACION, p_AnoAct, p_MesAct)
               
               l_int_NumEnt = 0: l_dbl_MtoDeu = 0
               Call fs_ObtieneDatosRCC(g_rst_Princi!HIPCIE_TDOCLI, g_rst_Princi!HIPCIE_NDOCLI, CStr(p_MesAnt), CStr(p_AnoAnt), l_int_NumEnt, l_dbl_MtoDeu)
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 25) = CDbl(Format(l_int_NumEnt, "#0"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 25), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 25)).NumberFormat = "##0"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 26) = CDbl(Format(l_dbl_MtoDeu, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 26), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 26)).NumberFormat = "###,###,##0.00"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 27) = fs_Busca_Excepciones(g_rst_Princi!OPERACION)
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 27), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 27)).NumberFormat = "#,##0"
               
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 28) = Trim(g_rst_Princi!NOMBRE_PROYECTO)
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 29) = g_rst_Princi!MICROEMPRESARIO
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 30) = g_rst_Princi!MATRIZ

               p_str_CadAux = ""
               p_str_CadAux = p_str_CadAux & "SELECT REGACC_FECCOM, TIPO_DESCRIPCION"
               p_str_CadAux = p_str_CadAux & "  FROM ("
               p_str_CadAux = p_str_CadAux & "        SELECT REGACC_FECACC, REGACC_HORACC, A.REGACC_FECCOM, TRIM(B.PARDES_DESCRI) AS TIPO_DESCRIPCION"
               p_str_CadAux = p_str_CadAux & "          FROM CBR_REGACC A"
               p_str_CadAux = p_str_CadAux & "          LEFT JOIN MNT_PARDES B ON B.PARDES_CODGRP = 308 AND B.PARDES_CODITE = A.REGACC_SITCOM"
               p_str_CadAux = p_str_CadAux & "         WHERE REGACC_TIPDOC = " & g_rst_Princi!HIPCIE_TDOCLI
               p_str_CadAux = p_str_CadAux & "           AND REGACC_NUMDOC = '" & Trim(g_rst_Princi!HIPCIE_NDOCLI) & "'"
               p_str_CadAux = p_str_CadAux & "           AND REGACC_TIPACC >= 101 AND REGACC_TIPACC <= 199"
               p_str_CadAux = p_str_CadAux & "         ORDER BY REGACC_FECACC DESC, REGACC_HORACC DESC) AA"
               p_str_CadAux = p_str_CadAux & "         WHERE AA.REGACC_FECCOM > 0 And ROWNUM = 1"

               If Not gf_EjecutaSQL(p_str_CadAux, r_rst_GenAux, 3) Then
                  Exit Sub
               End If
               
               If Not (r_rst_GenAux.BOF And r_rst_GenAux.EOF) Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 31) = "'" & gf_FormatoFecha(r_rst_GenAux!REGACC_FECCOM)
                  r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 32) = r_rst_GenAux!TIPO_DESCRIPCION
               Else
                  r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 29) = ""
                  r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 30) = ""
               End If
               
               r_rst_GenAux.Close
               Set r_rst_GenAux = Nothing
               
               'Clientes que han ingresado como Morosos
               '05/08/2014
               If CDbl(Format(g_rst_Princi!PROVISION2, "###,###,##0.00")) = 0 And (fs_MuestraTipo(g_rst_Princi!CLASIFICACION, g_rst_Princi!CLASIFICACION2) = "NUEVO") Then
                   r_dbl_CMoros = r_dbl_CMoros + CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00"))
               End If
               
               'Amortización, Cancelación, Regresan, Salen
               '05/08/2014
               If CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00")) < 1 And (fs_MuestraTipo(g_rst_Princi!CLASIFICACION, g_rst_Princi!CLASIFICACION2) = "MEJOR" Or fs_MuestraTipo(g_rst_Princi!CLASIFICACION, g_rst_Princi!CLASIFICACION2) = "IGUAL") Then
                   r_dbl_HipPrv = r_dbl_HipPrv + CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00"))
               End If
               
               'Clientes Morosos que han incrementado su provisión
               If CInt(g_rst_Princi!HIPCIE_CLACLI) = CInt(g_rst_Princi!HIPCIE_CLAALI) And CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00")) > 0 And CDbl(g_rst_Princi!PROVISION2) > 0 Then
                   r_dbl_CliMoI = r_dbl_CliMoI + CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00"))
               End If
               
               'Provisión Neta por Clientes Alineados
               If CInt(g_rst_Princi!HIPCIE_CLACLI) <> CInt(g_rst_Princi!HIPCIE_CLAALI) And CInt(g_rst_Princi!HIPCIE_CLAPRV) > 1 Then
                   r_dbl_PrvNet = r_dbl_PrvNet + CDbl(Format(g_rst_Princi!AJUSTE, "###,###,##0.00"))
               End If
               
               'Clientes Normales (Suma de Provision Voluntaria)
               If g_rst_Princi!CLASIFICACION = "NOR" Then
                  If g_rst_Princi!PROV_VOLUNT1 <> g_rst_Princi!PROV_VOLUNT2 Then
                     r_dbl_SumNor = r_dbl_SumNor + (g_rst_Princi!PROV_VOLUNT1 - g_rst_Princi!PROV_VOLUNT2)
                  End If
               End If
               
               r_dbl_SumCapBal = r_dbl_SumCapBal + CDbl(Format(g_rst_Princi!CAPITAL_BAL, "###,###,##0.00"))
               r_dbl_TotCapBal = r_dbl_TotCapBal + CDbl(Format(g_rst_Princi!CAPITAL_BAL, "###,###,##0.00"))
               r_dbl_SumCapAju = r_dbl_SumCapAju + CDbl(Format(g_rst_Princi!CAPITAL, "###,###,##0.00"))
               r_dbl_TotCapAju = r_dbl_TotCapAju + CDbl(Format(g_rst_Princi!CAPITAL, "###,###,##0.00"))
               r_dbl_SumPrv = r_dbl_SumPrv + CDbl(Format(g_rst_Princi!PROVISION, "###,###,##0.00"))
               r_dbl_TotPrv = r_dbl_TotPrv + CDbl(Format(g_rst_Princi!PROVISION, "###,###,##0.00"))
               r_dbl_SumVol = r_dbl_SumVol + CDbl(Format(g_rst_Princi!PROV_VOLUNT1, "###,###,##0.00"))
               r_dbl_TotVol = r_dbl_TotVol + CDbl(Format(g_rst_Princi!PROV_VOLUNT1, "###,###,##0.00"))
               
               r_int_NroFil = r_int_NroFil + 1
               g_rst_Princi.MoveNext
               DoEvents
            Loop
            
            If r_dbl_SumCapAju = 0 Then
               r_int_NroFil = r_int_NroFil - 1
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10) = CDbl(Format(r_dbl_SumCapBal, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10)).NumberFormat = "###,###,##0.00"
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11) = CDbl(Format(r_dbl_SumCapAju, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11)).NumberFormat = "###,###,##0.00"
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13) = CDbl(Format(r_dbl_SumPrv, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13)).NumberFormat = "###,###,##0.00"
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14) = CDbl(Format(r_dbl_SumVol, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14)).NumberFormat = "###,###,##0.00"
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10).Font.Bold = True
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11).Font.Bold = True
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13).Font.Bold = True
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14).Font.Bold = True
            End If
            r_int_nroaux = r_int_nroaux + 1
         Loop
        
         If Not r_bol_FlagOp Then r_int_NroFil = r_int_NroFil + 2
         r_bol_FlagOp = False
         r_int_Nindex = r_int_Nindex + 1
      Loop
    
      'Total General
      r_int_NroFil = r_int_NroFil + 1
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10) = CDbl(Format(r_dbl_TotCapBal, "###,###,##0.00"))
      r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10)).NumberFormat = "###,###,##0.00"
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11) = CDbl(Format(r_dbl_TotCapAju, "###,###,##0.00"))
      r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11)).NumberFormat = "###,###,##0.00"
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13) = CDbl(Format(r_dbl_TotPrv, "###,###,##0.00"))
      r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13)).NumberFormat = "###,###,##0.00"
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14) = CDbl(Format(r_dbl_TotVol, "###,###,##0.00"))
      r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14)).NumberFormat = "###,###,##0.00"
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10).Font.Bold = True
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11).Font.Bold = True
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13).Font.Bold = True
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14).Font.Bold = True
   End With
   
   'Cuadros
   r_int_NroFil = r_int_NroFil + 3
   With r_obj_Excel.Sheets(1)
      '*******
      'RESUMEN
      '*******
      
      'Titulo
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Font.Bold = True
      
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 3) = "RESUMEN:"
      r_int_NroFil = r_int_NroFil + 1
      
      '***********************
      'I. Creditos Comerciales
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 3) = "Créditos Comerciales"
      Call gf_DesAmoCreCom(l_str_PerMes, l_str_PerAno, l_str_MesAnt, l_str_AnoAnt, r_dbl_NvDCom, r_dbl_xAmCom)
      r_int_NroFil = r_int_NroFil + 1
      
      '1. Nuevos Desembolsos comerciales
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Nuevos Desembolsos"
      'r_dbl_NvDCom = gf_NvoCreCom(l_str_PerMes, l_str_PerAno)
      .Cells(r_int_NroFil, 5) = r_dbl_NvDCom
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 1
      
      '2. Por Amortizacion comerciales
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Por Amortización Comerciales"
      'r_dbl_xAmCom = gf_AmoCreCom(l_str_PerMes, l_str_PerAno, l_str_MesAnt, l_str_AnoAnt)
      .Cells(r_int_NroFil, 5) = r_dbl_xAmCom
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 1
      
      '*************************
      'II. Creditos Hipotecarios
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 3) = "Créditos Hipotecarios"
      r_int_NroFil = r_int_NroFil + 1
      
      '1. Nuevos Desembolsos hipotecarios
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Nuevos Desembolsos"
      r_dbl_NvDHip = gf_QNvoDH(l_str_PerMes, l_str_PerAno)
      .Cells(r_int_NroFil, 5) = r_dbl_NvDHip
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 1
      
      '2-a. Por Amortizacion hipotecarios
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Amortización, Cancelación, Regresan, Salen (transferidos)"
      r_dbl_PorHip1 = gf_PorHip(l_str_PerMes, l_str_PerAno, l_str_MesAnt, l_str_AnoAnt, 1)
      .Cells(r_int_NroFil, 5) = r_dbl_PorHip1
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 1
      
      '2-b. Por Amortizacion hipotecarios
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Amortización, Cancelación, Regresan, Salen"
      r_dbl_PorHip2 = gf_PorHip(l_str_PerMes, l_str_PerAno, l_str_MesAnt, l_str_AnoAnt, 2)
      .Cells(r_int_NroFil, 5) = r_dbl_PorHip2
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 1
      
      '3. Clientes ingresados como morosos
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Clientes que han ingresado como Morosos"
      .Cells(r_int_NroFil, 5) = r_dbl_CMoros
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 1
      
      '4. Clientes morosos que revierten provision
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Clientes Morosos que revierten Provisión"
      .Cells(r_int_NroFil, 5) = r_dbl_HipPrv
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 1
      
      '5. Clientes morosos que incrementan su provision
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Clientes Morosos que han incrementado su provisión"
      .Cells(r_int_NroFil, 5) = r_dbl_CliMoI
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 1
      
      '6. Clientes Morosos que ya estan al dia
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Clientes Morosos que ya están al día"
      r_dbl_CMoDia = gf_SalFav(l_str_PerMes, l_str_PerAno, l_str_MesAnt, l_str_AnoAnt)
      .Cells(r_int_NroFil, 5) = r_dbl_CMoDia
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 1
      
      '7. Provision Neta por clientes Alienados
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Provisión Neta por Clientes Alineados"
      .Cells(r_int_NroFil, 5) = r_dbl_PrvNet
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 1
      
      '8. Clientes Normales
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).Merge
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 4)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Cells(r_int_NroFil, 3) = "      Provisión Voluntaria"
      r_dbl_SumNor = gf_PrvVol(l_str_PerMes, l_str_PerAno, l_str_MesAnt, l_str_AnoAnt)
      .Cells(r_int_NroFil, 5) = r_dbl_SumNor
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      
      r_int_NroFil = r_int_NroFil + 1
      .Cells(r_int_NroFil, 5).Font.Bold = True
      .Cells(r_int_NroFil, 5) = r_dbl_NvDHip + r_dbl_NvDCom + r_dbl_PorHip1 + r_dbl_PorHip2 + r_dbl_xAmCom + r_dbl_CMoros + r_dbl_HipPrv + r_dbl_CliMoI + r_dbl_CMoDia + r_dbl_PrvNet + r_dbl_SumNor
      .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
      
      '***********************************************
      'DETALLE DE CLIENTES MOROSOS QUE YA ESTAN AL DIA
      '***********************************************
      r_dbl_TotPrv = 0
      r_int_NroFil = r_int_NroFil + 3
      .Cells(r_int_NroFil, 3) = "DETALLE DE CLIENTES MOROSOS Y ALINEADOS QUE YA ESTAN AL DIA"
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil, 5)).Merge
      .Cells(r_int_NroFil, 3).Font.Bold = True
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = r_int_NroFil + 1
      .Cells(r_int_NroFil, 3) = "OPERACION"
      .Cells(r_int_NroFil, 3).Font.Bold = True
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 4) = "NOMBRE DEL CLIENTE"
      .Cells(r_int_NroFil, 4).Font.Bold = True
      .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 5) = "MONTO"
      .Cells(r_int_NroFil, 5).Font.Bold = True
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignCenter
      
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & "SELECT HIPCIE_NUMOPE AS OPERACION, "
      r_str_Cadena = r_str_Cadena & "       TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS CLIENTE, "
      r_str_Cadena = r_str_Cadena & "       ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*(SELECT DISTINCT HIPCIE_TIPCAM FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & l_str_PerMes & " AND HIPCIE_PERANO = " & l_str_PerAno & ")),2)* -1 AS MONTO, "
      r_str_Cadena = r_str_Cadena & "       HIPCIE_DIAMOR AS MOROSIDAD "
      r_str_Cadena = r_str_Cadena & "  FROM CRE_HIPCIE "
      r_str_Cadena = r_str_Cadena & " INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPCIE_TDOCLI AND DATGEN_NUMDOC = HIPCIE_NDOCLI "
      r_str_Cadena = r_str_Cadena & " WHERE HIPCIE_PERANO = " & l_str_AnoAnt & " AND HIPCIE_PERMES = " & l_str_MesAnt & " AND HIPCIE_CLAPRV <> 0 "
      r_str_Cadena = r_str_Cadena & "   AND HIPCIE_NUMOPE NOT IN "
      r_str_Cadena = r_str_Cadena & "          (SELECT HIPCIE_NUMOPE "
      r_str_Cadena = r_str_Cadena & "             FROM CRE_HIPCIE "
      r_str_Cadena = r_str_Cadena & "            WHERE HIPCIE_PERANO = " & l_str_PerAno & " AND HIPCIE_PERMES = " & l_str_PerMes & " AND HIPCIE_CLAPRV <> 0) "
      r_str_Cadena = r_str_Cadena & " ORDER BY MOROSIDAD DESC "
      
      If Not gf_EjecutaSQL(r_str_Cadena, g_rst_GenAux, 3) Then
         Exit Sub
      End If
        
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         r_bol_Flag1 = True
         r_bol_Flag2 = True
         g_rst_GenAux.MoveFirst
         
         Do While Not g_rst_GenAux.EOF
            If r_bol_Flag1 = True Then
               If g_rst_GenAux!MOROSIDAD > 30 Then
                  r_int_NroFil = r_int_NroFil + 1
                  .Cells(r_int_NroFil, 3) = "POR COBRANZAS"
                  .Cells(r_int_NroFil, 3).Font.Bold = True
                  .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignLeft
                  r_bol_Flag1 = False
               End If
            End If
            If r_bol_Flag2 = True Then
               If g_rst_GenAux!MOROSIDAD < 31 Then
                  r_int_NroFil = r_int_NroFil + 1
                  .Cells(r_int_NroFil, 3) = "POR ALINEAMIENTO"
                  .Cells(r_int_NroFil, 3).Font.Bold = True
                  .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignLeft
                  r_bol_Flag2 = False
               End If
            End If
            
            r_int_NroFil = r_int_NroFil + 1
            .Cells(r_int_NroFil, 3) = g_rst_GenAux!OPERACION
            .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
            .Cells(r_int_NroFil, 4) = g_rst_GenAux!CLIENTE
            .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 5) = g_rst_GenAux!MONTO
            .Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil, 5)).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignRight
            r_dbl_TotPrv = r_dbl_TotPrv + g_rst_GenAux!MONTO
            
            g_rst_GenAux.MoveNext
         Loop
         
         r_int_NroFil = r_int_NroFil + 1
         .Cells(r_int_NroFil, 5).Font.Bold = True
         .Cells(r_int_NroFil, 5) = Format(r_dbl_TotPrv, "###,###,##0.00")
      End If
        
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
    End With
    
   '***********************************************
   'RESUMEN POR PERIODO DE CLIENTES SEGÚN TIPO
   '***********************************************
   r_int_NumMes = 1
   Do While r_int_NumMes <= CInt(p_MesAct)
      p_MesActN = r_int_NumMes
      p_AnoActN = ipp_PerAno.Text
      p_MesAntN = IIf(r_int_NumMes = 1, 12, r_int_NumMes - 1)
      p_AnoAntN = IIf(r_int_NumMes = 1, CLng(ipp_PerAno.Text) - 1, ipp_PerAno.Text)
      
      g_str_Parame = "USP_RPT_CLI_MORO_ALIN ("
      g_str_Parame = g_str_Parame & CInt(p_MesActN) & ", "
      g_str_Parame = g_str_Parame & CInt(p_AnoActN) & ", "
      g_str_Parame = g_str_Parame & CInt(p_MesAntN) & ", "
      g_str_Parame = g_str_Parame & CInt(p_AnoAntN) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & "REPORTE " & UCase(SSPanel2.Caption) & "',1) "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         MsgBox "Error al ejecutar el Procedimiento USP_CUR_GEN_CLI_MORO_ALIN.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If

      If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
         g_rst_GenAux.Close
         Set g_rst_GenAux = Nothing
         MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         g_rst_GenAux.MoveFirst
         Do While Not g_rst_GenAux.EOF
            g_rst_Listas.AddNew
            g_rst_Listas.Fields(0).Value = g_rst_GenAux!RPT_PERMES
            g_rst_Listas.Fields(1).Value = g_rst_GenAux!RPT_PERANO
            g_rst_Listas.Fields(2).Value = g_rst_GenAux!RPT_DESCRI
            g_rst_Listas.Fields(3).Value = g_rst_GenAux!RPT_VALNUM01
            
            g_rst_Listas.Update
            g_rst_GenAux.MoveNext
         Loop
      End If
      r_int_NumMes = r_int_NumMes + 1
   Loop
   
   r_obj_Excel.Sheets(2).Name = "PERIODOS"
   r_int_NroFil = 2
   
   With r_obj_Excel.Sheets(2)
      .Range("A" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("A" & r_int_NroFil & ":M" & r_int_NroFil & "").Font.Bold = True
      .Range("A" & r_int_NroFil & ":M" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_NroFil & ":M" & r_int_NroFil & "").Font.Size = 14
      .Range("A" & r_int_NroFil & ":M" & r_int_NroFil & "").Font.Underline = xlUnderlineStyleSingle
      
      .Cells(r_int_NroFil, 1) = "RESUMEN DE PROVISIONES AL MES DE " & Trim(cmb_PerMes.Text) & " DEL " & Trim(ipp_PerAno.Text)
      r_int_NroFil = r_int_NroFil + 2
      
      r_int_NroFil = r_int_NroFil + 1
      .Cells(r_int_NroFil, 1) = "RESUMEN"
      .Cells(r_int_NroFil, 2) = "ENERO"
      .Cells(r_int_NroFil, 3) = "FEBRERO"
      .Cells(r_int_NroFil, 4) = "MARZO"
      .Cells(r_int_NroFil, 5) = "ABRIL"
      .Cells(r_int_NroFil, 6) = "MAYO"
      .Cells(r_int_NroFil, 7) = "JUNIO"
      .Cells(r_int_NroFil, 8) = "JULIO"
      .Cells(r_int_NroFil, 9) = "AGOSTO"
      .Cells(r_int_NroFil, 10) = "SETIEMBRE"
      .Cells(r_int_NroFil, 11) = "OCTUBRE"
      .Cells(r_int_NroFil, 12) = "NOVIEMBRE"
      .Cells(r_int_NroFil, 13) = "DICIEMBRE"
      
      .Cells(6, 1) = "ALINEADO"
      .Cells(7, 1) = "ALINEADO-MOROSO"
      .Cells(8, 1) = "IGUAL"
      .Cells(9, 1) = "MEJOR"
      .Cells(10, 1) = "NUEVO"
      .Cells(11, 1) = "PEOR"
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 13)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 13)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 20
      .Columns("B").ColumnWidth = 14
      .Columns("C").ColumnWidth = 14
      .Columns("D").ColumnWidth = 14
      .Columns("E").ColumnWidth = 14
      .Columns("F").ColumnWidth = 14
      .Columns("G").ColumnWidth = 14
      .Columns("H").ColumnWidth = 14
      .Columns("I").ColumnWidth = 14
      .Columns("J").ColumnWidth = 14
      .Columns("K").ColumnWidth = 14
      .Columns("L").ColumnWidth = 14
      .Columns("M").ColumnWidth = 14
      
      .Columns("A").HorizontalAlignment = xlHAlignLeft
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      
      If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
         g_rst_Listas.MoveFirst
         Do While Not g_rst_Listas.EOF
            r_int_NroFil = r_int_NroFil + 1

INGRE:
            If .Cells(r_int_NroFil, 1) = Trim(g_rst_Listas!TIPO) Then
               .Cells(r_int_NroFil, g_rst_Listas!Mes + 1) = Trim(g_rst_Listas!CANTIDAD)
            ElseIf .Cells(r_int_NroFil, 1) = "" Then
               r_int_NroFil = 6
               GoTo INGRE
            ElseIf IsNull(Trim(g_rst_Listas!TIPO)) Then
               GoTo SALTO1
            Else
               .Cells(r_int_NroFil, g_rst_Listas!Mes + 1) = 0
               r_int_NroFil = r_int_NroFil + 1
               GoTo INGRE
            End If
                
SALTO1:
            g_rst_Listas.MoveNext
         Loop
      End If
      
      .Range("A" & 12 & ":M" & 12 & "").Font.Bold = True
      For r_int_NroFil = 2 To p_MesAct + 1
         .Cells(12, r_int_NroFil) = "=SUM(R[-6]C:R[-1]C)"
      Next
   End With
    
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Function fs_Obtiene_TasaProvision(ByVal p_NumOpe As String, ByVal p_Clacli As Integer, ByVal p_TipGar As Integer, ByVal p_Tasa As Double) As Double
Dim r_str_Parame     As String
Dim r_int_TipGar     As Integer
Dim r_rst_ClaTim     As ADODB.Recordset
   
   r_int_TipGar = 0
   fs_Obtiene_TasaProvision = p_Tasa
   
   '********** DETERMINA SI TIENE CLASICACION DUDOSA POR MAS DE 36 MESES **********
   If p_Clacli = 3 Then
      'r_str_Parame = ""
      'r_str_Parame = r_str_Parame & "SELECT DISTINCT HIPCIE_CLAPRV, COUNT(*) AS CONTADOR "
      'r_str_Parame = r_str_Parame & "  FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLAPRV "
      'r_str_Parame = r_str_Parame & "          FROM CRE_HIPCIE "
      'r_str_Parame = r_str_Parame & "         WHERE HIPCIE_PERMES > 0 "
      'r_str_Parame = r_str_Parame & "           AND HIPCIE_PERANO > 2010 "
      'r_str_Parame = r_str_Parame & "           AND HIPCIE_NUMOPE = '" & p_NumOpe & "' "
      'r_str_Parame = r_str_Parame & "           AND HIPCIE_CLAPRV = 3 "
      'r_str_Parame = r_str_Parame & "         ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
      'r_str_Parame = r_str_Parame & " WHERE ROWNUM < 37 "
      'r_str_Parame = r_str_Parame & " GROUP BY HIPCIE_CLAPRV "
      
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT COUNT(*) AS CONTADOR "
      r_str_Parame = r_str_Parame & "  FROM (SELECT HIPCIE_CLACLI "
      r_str_Parame = r_str_Parame & "          FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLACLI "
      r_str_Parame = r_str_Parame & "                  FROM CRE_HIPCIE "
      r_str_Parame = r_str_Parame & "                 WHERE HIPCIE_PERMES > 0 "
      r_str_Parame = r_str_Parame & "                   AND HIPCIE_PERANO > 2014 "
      r_str_Parame = r_str_Parame & "                   AND HIPCIE_NUMOPE = '" & p_NumOpe & "' "
      r_str_Parame = r_str_Parame & "                 ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
      r_str_Parame = r_str_Parame & "         WHERE ROWNUM < 37) "
      r_str_Parame = r_str_Parame & " WHERE HIPCIE_CLACLI = 3 "
      
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_ClaTim, 3) Then
         Exit Function
      End If
      
      If Not (r_rst_ClaTim.BOF And r_rst_ClaTim.EOF) Then
         r_rst_ClaTim.MoveFirst
         If r_rst_ClaTim!CONTADOR = 36 Then
            r_int_TipGar = 5
         End If
      End If
      
      r_rst_ClaTim.Close
      Set r_rst_ClaTim = Nothing
   End If
   
   '********** DETERMINA SI TIENE CLASICACION PERDIDA POR MAS DE 24 MESES **********
   If p_Clacli = 4 Then
      'r_str_Parame = ""
      'r_str_Parame = r_str_Parame & "SELECT DISTINCT HIPCIE_CLAPRV, COUNT(*) AS CONTADOR "
      'r_str_Parame = r_str_Parame & "  FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLAPRV "
      'r_str_Parame = r_str_Parame & "          FROM CRE_HIPCIE "
      'r_str_Parame = r_str_Parame & "         WHERE HIPCIE_PERMES > 0 "
      'r_str_Parame = r_str_Parame & "           AND HIPCIE_PERANO > 2010 "
      'r_str_Parame = r_str_Parame & "           AND HIPCIE_NUMOPE = '" & p_NumOpe & "' "
      'r_str_Parame = r_str_Parame & "           AND HIPCIE_CLAPRV = 4 "
      'r_str_Parame = r_str_Parame & "         ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
      'r_str_Parame = r_str_Parame & " WHERE ROWNUM < 25 "
      'r_str_Parame = r_str_Parame & " GROUP BY HIPCIE_CLAPRV "
      
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT COUNT(*) AS CONTADOR "
      r_str_Parame = r_str_Parame & "  FROM (SELECT HIPCIE_CLACLI "
      r_str_Parame = r_str_Parame & "          FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLACLI "
      r_str_Parame = r_str_Parame & "                  FROM CRE_HIPCIE "
      r_str_Parame = r_str_Parame & "                 WHERE HIPCIE_PERMES > 0 "
      r_str_Parame = r_str_Parame & "                   AND HIPCIE_PERANO > 2015 "
      r_str_Parame = r_str_Parame & "                   AND HIPCIE_NUMOPE = '" & p_NumOpe & "' "
      r_str_Parame = r_str_Parame & "                 ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
      r_str_Parame = r_str_Parame & "         WHERE ROWNUM < 25) "
      r_str_Parame = r_str_Parame & " WHERE HIPCIE_CLACLI = 4 "
      
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_ClaTim, 3) Then
         Exit Function
      End If
      
      If Not (r_rst_ClaTim.BOF And r_rst_ClaTim.EOF) Then
         r_rst_ClaTim.MoveFirst
         If r_rst_ClaTim!CONTADOR = 24 Then
            r_int_TipGar = 5
         End If
      End If
      
      r_rst_ClaTim.Close
      Set r_rst_ClaTim = Nothing
   End If
   
   'CAMBIA DE CLASIFICACION
   If r_int_TipGar = 5 Then
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT TIPPRV_PORCEN AS TASA"
      r_str_Parame = r_str_Parame & "  FROM CTB_TIPPRV"
      r_str_Parame = r_str_Parame & " WHERE TIPPRV_TIPPRV = '2' "
      r_str_Parame = r_str_Parame & "   AND TIPPRV_CLACRE = '13'"
      r_str_Parame = r_str_Parame & "   AND TIPPRV_CLFCRE = " & p_Clacli & " "
      r_str_Parame = r_str_Parame & "   AND TIPPRV_CLAGAR = 1"
      
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_ClaTim, 3) Then
         Exit Function
      End If
    
      If Not (r_rst_ClaTim.BOF And r_rst_ClaTim.EOF) Then
         r_rst_ClaTim.MoveFirst
         fs_Obtiene_TasaProvision = r_rst_ClaTim!TASA
      End If
      
      r_rst_ClaTim.Close
      Set r_rst_ClaTim = Nothing
   End If
End Function

Private Sub fs_GenExc_Actual_Nuevo()
Dim r_rst_PrvAnt     As ADODB.Recordset
Dim r_rst_PrvAct     As ADODB.Recordset
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_ClaInt     As Integer
Dim r_bol_Muestra    As Boolean
Dim r_str_TipDoc     As String
Dim r_str_NumDoc     As String
Dim r_str_ClaInt     As String
Dim r_int_ClaPrv     As Integer
Dim r_str_ClaPrv     As String
Dim r_dbl_PrvAnt     As Double
Dim r_dbl_MtoPrv     As Double
Dim r_dbl_TasPrv     As Double
Dim r_int_TipGar     As Integer
Dim r_int_Refina     As Integer
Dim r_dbl_Capita     As Double
   
   '****************************
   'determina ultimo rcc cargado
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM (SELECT DISTINCT RCCCAB_PERANO, RCCCAB_PERMES "
   g_str_Parame = g_str_Parame & "          FROM CLI_RCCCAB "
   g_str_Parame = g_str_Parame & "         ORDER BY RCCCAB_PERANO DESC, RCCCAB_PERMES DESC) "
   g_str_Parame = g_str_Parame & " WHERE ROWNUM < 2 "
   g_str_Parame = g_str_Parame & " ORDER BY RCCCAB_PERANO DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   l_str_MesRcc = g_rst_Princi!RCCCAB_PERMES
   l_str_AnoRcc = g_rst_Princi!RCCCAB_PERANO
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   '*********************************
   'determina ultimo cierre procesado
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM (SELECT DISTINCT HIPCIE_PERANO, HIPCIE_PERMES "
   g_str_Parame = g_str_Parame & "          FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & "         ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
   g_str_Parame = g_str_Parame & " WHERE ROWNUM < 2 "
   g_str_Parame = g_str_Parame & " ORDER BY HIPCIE_PERANO DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   l_str_PerMes = g_rst_Princi!HIPCIE_PERMES
   l_str_PerAno = g_rst_Princi!HIPCIE_PERANO
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   '******************************************
   'consulta de la provision al dia de proceso
   g_str_Parame = gf_Query_PrvAct(l_str_PerMes, l_str_PerAno)
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_PrvAct, 3) Then
      Exit Sub
   End If
   
   If r_rst_PrvAct.BOF And r_rst_PrvAct.EOF Then
      r_rst_PrvAct.Close
      Set r_rst_PrvAct = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   '******************************************
   'consulta de la provision del ultimo cierre
   g_str_Parame = gf_Query_PrvAnt(l_str_PerMes, l_str_PerAno)
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_PrvAnt, 3) Then
      Exit Sub
   End If
   
   If r_rst_PrvAnt.BOF And r_rst_PrvAnt.EOF Then
      r_rst_PrvAnt.Close
      Set r_rst_PrvAnt = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   '*******************************
   'configuracion inicial del excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 2
   
   With r_obj_Excel.ActiveSheet
      .Range("A" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Range("A" & r_int_NroFil & ":P" & r_int_NroFil & "").Font.Bold = True
      .Range("A" & r_int_NroFil & ":P" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_NroFil & ":P" & r_int_NroFil & "").Font.Size = 14
      .Range("A" & r_int_NroFil & ":P" & r_int_NroFil & "").Font.Underline = xlUnderlineStyleSingle
      
      .Cells(r_int_NroFil, 1) = "PROVISIÓN DE CLIENTES MOROSOS Y ALINEADOS AL " & Day(date) & " DE " & UCase(MonthName(Month(date))) & " DEL " & Year(date)
      r_int_NroFil = r_int_NroFil + 2
      
      .Range("J" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("M" & r_int_NroFil & ":O" & r_int_NroFil & "").Merge
      .Range("J" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("M" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("J" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
      .Range("M" & r_int_NroFil & ":O" & r_int_NroFil & "").Font.Bold = True
      
      .Cells(r_int_NroFil, 10) = "ACTUAL"
      .Cells(r_int_NroFil, 13) = "ANTERIOR"
      
      r_int_NroFil = r_int_NroFil + 1
      .Cells(r_int_NroFil, 1) = "PRODUCTO"
      .Cells(r_int_NroFil, 2) = "TIPO"
      .Cells(r_int_NroFil, 3) = "NOMBRE DEL CLIENTE"
      .Cells(r_int_NroFil, 4) = "ATRASO REAL"
      .Cells(r_int_NroFil, 5) = "ATRASO PROYEC."
      .Cells(r_int_NroFil, 6) = "GARANTÍA"
      .Cells(r_int_NroFil, 7) = "VAL. GARANTÍA"
      .Cells(r_int_NroFil, 8) = "CAPITAL"
      .Cells(r_int_NroFil, 9) = "TASA %"
      .Cells(r_int_NroFil, 10) = "PROVISIÓN"
      .Cells(r_int_NroFil, 11) = "CLA. ALINEADA"
      .Cells(r_int_NroFil, 12) = "CLA. INTERNA"
      .Cells(r_int_NroFil, 13) = "PROVISIÓN"
      .Cells(r_int_NroFil, 14) = "CLA. ALINEADA"
      .Cells(r_int_NroFil, 15) = "CLA. INTERNA"
      .Cells(r_int_NroFil, 16) = "AJUSTE"
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 16)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 16)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 11
      .Columns("B").ColumnWidth = 12
      .Columns("C").ColumnWidth = 46
      .Columns("D").ColumnWidth = 13
      .Columns("E").ColumnWidth = 16
      .Columns("F").ColumnWidth = 11
      .Columns("G").ColumnWidth = 15
      .Columns("H").ColumnWidth = 12
      .Columns("I").ColumnWidth = 10
      .Columns("J").ColumnWidth = 15
      .Columns("K").ColumnWidth = 14
      .Columns("L").ColumnWidth = 13
      .Columns("M").ColumnWidth = 12
      .Columns("N").ColumnWidth = 14
      .Columns("O").ColumnWidth = 13
      .Columns("P").ColumnWidth = 11
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").HorizontalAlignment = xlHAlignCenter
   End With
   
   '*****************************
   'muestra informacion en grilla
   r_int_NroFil = r_int_NroFil + 1
   Do While Not r_rst_PrvAct.EOF
      r_str_ClaInt = ""
      r_str_TipDoc = r_rst_PrvAct!TIPO_DOC
      r_str_NumDoc = r_rst_PrvAct!NUMERO_DOC
      r_int_ClaInt = r_rst_PrvAct!CLA_INT
      r_str_ClaPrv = ""
      r_int_ClaPrv = 0
      r_dbl_MtoPrv = 0
      r_dbl_TasPrv = 0
      r_int_TipGar = r_rst_PrvAct!TIP_GARANT
      r_int_Refina = r_rst_PrvAct!REFINANCIADO
      r_dbl_Capita = r_rst_PrvAct!CAPITAL_SOL
      
      Call fs_Obtiene_ClasificacionAlineada(r_str_TipDoc, r_str_NumDoc, r_int_ClaInt, r_str_ClaPrv, r_int_ClaPrv, r_dbl_TasPrv, r_int_TipGar, r_int_Refina, r_dbl_Capita)
      
      r_bol_Muestra = True
      If (r_int_ClaInt = r_int_ClaPrv) And (r_int_ClaInt = 0) Then
         r_bol_Muestra = False
      End If
      
      If r_bol_Muestra Then
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 1) = Trim(r_rst_PrvAct!PRODUCTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 2) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 3) = Trim(r_rst_PrvAct!NOMBRE_CLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 4) = r_rst_PrvAct!DIAS_ATRA_REAL
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 4), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 4)).NumberFormat = "#,##0"
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5) = r_rst_PrvAct!DIAS_ATRA_CALC
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5)).NumberFormat = "#,##0"
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 6) = Trim(r_rst_PrvAct!TIPO_GARANTIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 7) = CDbl(Format(r_rst_PrvAct!VALOR_GARANTIA, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 7), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 7)).NumberFormat = "###,###,##0.00"
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 8) = CDbl(Format(r_rst_PrvAct!CAPITAL_SOL, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 8), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 8)).NumberFormat = "###,###,##0.00"
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 9) = CDbl(Format(r_dbl_TasPrv, "##0.00"))
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 9), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 9)).NumberFormat = "##0.00"
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 11) = Trim(r_str_ClaPrv)
         
         Select Case r_int_ClaInt
            Case 0: r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 12) = "NOR": r_str_ClaInt = "NOR"
            Case 1: r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 12) = "CPP": r_str_ClaInt = "CPP"
            Case 2: r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 12) = "DEF": r_str_ClaInt = "DEF"
            Case 3: r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 12) = "DUD": r_str_ClaInt = "DUD"
            Case 4: r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 12) = "PER": r_str_ClaInt = "PER"
         End Select
         
         'Datos de la provision anterior
         r_dbl_PrvAnt = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14) = ""
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 15) = ""
         r_rst_PrvAnt.MoveFirst
         Do While Not r_rst_PrvAnt.EOF
            If r_rst_PrvAct!OPERACION = r_rst_PrvAnt!OPERACION2 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13) = CDbl(Format(r_rst_PrvAnt!PROVISION2, "###,###,##0.00"))
               r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 13)).NumberFormat = "###,###,##0.00"
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14) = Trim(r_rst_PrvAnt!CLA_ALI2)
               r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 15) = Trim(r_rst_PrvAnt!CLA_INT2)
               r_dbl_PrvAnt = r_rst_PrvAnt!PROVISION2
               Exit Do
            End If
            r_rst_PrvAnt.MoveNext
         Loop
         
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 2) = fs_MuestraTipo(Trim(r_str_ClaPrv), Trim(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 14)))
         Call fs_CalculaProvision(r_rst_PrvAct!OPERACION, r_rst_PrvAct!COD_PRODUC, r_rst_PrvAct!COD_SUBPROD, r_rst_PrvAct!CAPITAL_SOL, r_rst_PrvAct!SALDO_TNC, r_rst_PrvAct!SALDO_TC, r_rst_PrvAct!TIPO_MONEDA, r_rst_PrvAct!TIPO_CAMBIO, r_rst_PrvAct!TIP_GARANT, r_rst_PrvAct!MONTO_GARANTIA, r_rst_PrvAct!MONEDA_GARANTIA, r_rst_PrvAct!FECHA_DESEMBOLSO, r_rst_PrvAct!CAPITAL_VENCIDO, r_int_ClaPrv, r_dbl_MtoPrv)
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10) = CDbl(Format(r_dbl_MtoPrv, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10)).NumberFormat = "###,###,##0.00"
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 16) = CDbl(Format(r_dbl_MtoPrv - r_dbl_PrvAnt, "###,###,##0.00"))
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 16), r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 16)).NumberFormat = "###,###,##0.00"
         
         r_int_NroFil = r_int_NroFil + 1
      End If
      
      r_rst_PrvAct.MoveNext
      DoEvents
   Loop
   
   r_rst_PrvAct.Close
   Set r_rst_PrvAct = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_CalculaProvision(ByVal p_NumOpe As String, ByVal p_CodPrd As String, ByVal p_CodSub As String, ByVal p_SalTot As Double, ByVal p_SalCap As Double, ByVal p_SalCon As Double, ByVal p_TipMon As Integer, ByVal p_TipCam As Double, ByVal p_TipGar As Integer, ByVal p_MtoGar As Double, ByVal p_MonGar As Integer, ByVal p_FecDes As String, ByVal p_CapVen As Double, ByVal p_ClaPrv As Integer, ByRef p_PrvEsp As Double)
Dim r_arr_DetGar()      As modprc_g_tpo_DetGar
Dim r_arr_TipPrv()      As modprc_g_tpo_TipPrv
Dim r_rst_Tablas        As ADODB.Recordset
Dim r_str_Cadena        As String
Dim r_int_Contad        As Integer
Dim r_int_ClaGar        As Integer
Dim r_dbl_GtoJud        As Double
Dim r_dbl_PerFmv        As Double
Dim r_dbl_MtoCga        As Double
Dim r_dbl_MtoSga        As Double
Dim r_dbl_Porce1        As Double
Dim r_dbl_Porce2        As Double

   'Tabla de provisiones
   r_str_Cadena = "SELECT * FROM CTB_TIPPRV WHERE TIPPRV_CLACRE = '13' "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Tablas, 3) Then
      r_rst_Tablas.Close
      Set r_rst_Tablas = Nothing
      Exit Sub
   End If
   
   r_rst_Tablas.MoveFirst
   ReDim r_arr_TipPrv(0)
   Do While Not r_rst_Tablas.EOF
      ReDim Preserve r_arr_TipPrv(UBound(r_arr_TipPrv) + 1)
      r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_TipPrv = CInt(r_rst_Tablas!TipPrv_TipPrv)
      r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_CodCla = CInt(r_rst_Tablas!TIPPRV_CLFCRE)
      r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_ClaGar = CInt(r_rst_Tablas!TipPrv_ClaGar)
      r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_Porcen = r_rst_Tablas!TipPrv_Porcen
      r_rst_Tablas.MoveNext
   Loop
   
   r_rst_Tablas.Close
   Set r_rst_Tablas = Nothing
   
   'Tabla de garantías
   r_str_Cadena = "SELECT * FROM CTB_DETGAR "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Tablas, 3) Then
      r_rst_Tablas.Close
      Set r_rst_Tablas = Nothing
      Exit Sub
   End If
   
   r_rst_Tablas.MoveFirst
   ReDim r_arr_DetGar(0)
   Do While Not r_rst_Tablas.EOF
      ReDim Preserve r_arr_DetGar(UBound(r_arr_DetGar) + 1)
      r_arr_DetGar(UBound(r_arr_DetGar)).DetGar_Codigo = CInt(r_rst_Tablas!DetGar_Codigo)
      r_arr_DetGar(UBound(r_arr_DetGar)).DetGar_ClaGar = CInt(r_rst_Tablas!DetGar_ClaGar)
      r_rst_Tablas.MoveNext
   Loop
   
   r_rst_Tablas.Close
   Set r_rst_Tablas = Nothing
   
   'Determinando Clase de Garantía según Tipo de Garantía
   r_int_ClaGar = 0
   For r_int_Contad = 1 To UBound(r_arr_DetGar)
      If r_arr_DetGar(r_int_Contad).DetGar_Codigo = p_TipGar Then
         r_int_ClaGar = r_arr_DetGar(r_int_Contad).DetGar_ClaGar
         Exit For
      End If
   Next r_int_Contad
   
   p_PrvEsp = 0
   r_dbl_PerFmv = 0
   r_dbl_GtoJud = 0
   r_dbl_MtoCga = 0
   r_dbl_MtoSga = 0
   
   'Determina gastos judiciales
   r_dbl_GtoJud = modprc_ff_CalculaGastosJudicial(p_CodPrd, p_CodSub, (p_SalCap + p_SalCon), p_TipMon, p_TipCam, p_TipGar, p_MtoGar, p_MonGar)
   
   'Determina montos base para el calculo de la provision
   Call modprc_ff_CalculaMontosBaseProv2(r_dbl_MtoCga, r_dbl_MtoSga, r_dbl_GtoJud, p_SalCap, p_SalCon, p_CodPrd, p_TipMon, p_TipCam, p_FecDes, p_TipGar, p_MtoGar, p_MonGar, 0, 1)
   
   r_dbl_MtoCga = Format(r_dbl_MtoCga, "###,##0.00")
   r_dbl_MtoSga = Format(r_dbl_MtoSga, "###,##0.00")
   
   'Si Clasificacion del deudor es diferente de normal: PROVISION ESPECIFICA
   If p_ClaPrv <> 0 Then
      r_dbl_Porce1 = (modprc_gf_PorcenProv(r_arr_TipPrv, 2, p_ClaPrv, 1) / 100)
      r_dbl_Porce2 = (modprc_gf_PorcenProv(r_arr_TipPrv, 2, p_ClaPrv, 2) / 100)
      p_PrvEsp = (r_dbl_Porce1 * r_dbl_MtoSga) + (r_dbl_Porce2 * r_dbl_MtoCga)
      
      '********** DETERMINA SI TIENE CLASICACION DUDOSA POR MAS DE 36 MESES **********
      If p_ClaPrv = 3 Then
         r_str_Cadena = ""
         r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS CONTADOR "
         r_str_Cadena = r_str_Cadena & "  FROM (SELECT HIPCIE_CLACLI "
         r_str_Cadena = r_str_Cadena & "          FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLACLI "
         r_str_Cadena = r_str_Cadena & "                  FROM CRE_HIPCIE "
         r_str_Cadena = r_str_Cadena & "                 WHERE HIPCIE_PERMES > 0 "
         r_str_Cadena = r_str_Cadena & "                   AND HIPCIE_PERANO > 2014 "
         r_str_Cadena = r_str_Cadena & "                   AND HIPCIE_NUMOPE = '" & p_NumOpe & "' "
         r_str_Cadena = r_str_Cadena & "                 ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
         r_str_Cadena = r_str_Cadena & "         WHERE ROWNUM < 37) "
         r_str_Cadena = r_str_Cadena & " WHERE HIPCIE_CLACLI = 3 "
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Tablas, 3) Then
            Exit Sub
         End If
         
         If Not (r_rst_Tablas.BOF And r_rst_Tablas.EOF) Then
            r_rst_Tablas.MoveFirst
            If r_rst_Tablas!CONTADOR = 36 Then
               p_PrvEsp = p_CapVen
            End If
         End If
         
         r_rst_Tablas.Close
         Set r_rst_Tablas = Nothing
      End If
   
      '********** DETERMINA SI TIENE CLASICACION PERDIDA POR MAS DE 24 MESES **********
      If p_ClaPrv = 4 Then
         r_str_Cadena = ""
         r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS CONTADOR "
         r_str_Cadena = r_str_Cadena & "  FROM (SELECT HIPCIE_CLACLI "
         r_str_Cadena = r_str_Cadena & "          FROM (SELECT HIPCIE_PERANO, HIPCIE_PERMES, HIPCIE_CLACLI "
         r_str_Cadena = r_str_Cadena & "                  FROM CRE_HIPCIE "
         r_str_Cadena = r_str_Cadena & "                 WHERE HIPCIE_PERMES > 0 "
         r_str_Cadena = r_str_Cadena & "                   AND HIPCIE_PERANO > 2014 "
         r_str_Cadena = r_str_Cadena & "                   AND HIPCIE_NUMOPE = '" & p_NumOpe & "' "
         r_str_Cadena = r_str_Cadena & "                 ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC) "
         r_str_Cadena = r_str_Cadena & "         WHERE ROWNUM < 25) "
         r_str_Cadena = r_str_Cadena & " WHERE HIPCIE_CLACLI = 4 "
         
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Tablas, 3) Then
            Exit Sub
         End If
         
         If Not (r_rst_Tablas.BOF And r_rst_Tablas.EOF) Then
            r_rst_Tablas.MoveFirst
            If r_rst_Tablas!CONTADOR = 24 Then
               p_PrvEsp = p_CapVen
            End If
         End If
      End If
   End If
   
   If p_TipMon <> 1 Then
      p_PrvEsp = Format(p_PrvEsp * p_TipCam, "###,###,##0.00")
   End If
End Sub

Private Sub fs_Obtiene_ClasificacionAlineada(ByVal p_TipDoc As String, ByVal p_NumDoc As String, ByVal p_ClaInt As String, ByRef p_str_ClaPrv As String, ByRef p_int_ClaPrv As Integer, ByRef p_TasPrv As Double, ByVal p_TipGar As Integer, ByVal p_Refina As Integer, ByVal p_Capita As Double)
Dim r_str_Cadena     As String
Dim r_dbl_TotDeu     As Double
Dim r_int_ClaInt     As Integer
Dim r_int_ClaAli     As Integer
Dim r_rst_DatRcc     As ADODB.Recordset
Dim r_rst_PrcPrv     As ADODB.Recordset
   
   'inicializa
   r_int_ClaAli = p_ClaInt
   r_int_ClaInt = p_ClaInt
   p_int_ClaPrv = p_ClaInt
   p_TasPrv = 0
   
   'obtiene datos del RCC del cliente
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT RCCDET_CODEMP, RCCDET_CLASIF, SUM(RCCDET_MTOSOL) AS DEUSOL, SUM(RCCDET_MTODOL) AS DEUDOL "
   r_str_Cadena = r_str_Cadena & "  FROM CLI_RCCDET "
   r_str_Cadena = r_str_Cadena & " WHERE RCCDET_TIPDOC = " & Trim(p_TipDoc) & " "
   r_str_Cadena = r_str_Cadena & "   AND RCCDET_NUMDOC = '" & Trim(p_NumDoc) & "' "
   r_str_Cadena = r_str_Cadena & "   AND RCCDET_PERMES = '" & CStr(l_str_MesRcc) & "' "
   r_str_Cadena = r_str_Cadena & "   AND RCCDET_PERANO = '" & CStr(l_str_AnoRcc) & "' "
   r_str_Cadena = r_str_Cadena & " GROUP BY RCCDET_CODEMP, RCCDET_CLASIF "
   r_str_Cadena = r_str_Cadena & " ORDER BY RCCDET_CODEMP ASC, RCCDET_CLASIF DESC "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_DatRcc, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_DatRcc.BOF And r_rst_DatRcc.EOF) Then
      'Sumariza las Deudas del cliente en el Sistema Financiero
      r_dbl_TotDeu = p_Capita
      r_rst_DatRcc.MoveFirst
      Do While Not r_rst_DatRcc.EOF
         r_dbl_TotDeu = r_dbl_TotDeu + r_rst_DatRcc!DEUSOL + r_rst_DatRcc!DEUDOL
         r_rst_DatRcc.MoveNext
      Loop
      
      'Obtiene la peor clasificacion reportada del cliente
      r_rst_DatRcc.MoveFirst
      Do While Not r_rst_DatRcc.EOF
         If r_rst_DatRcc!RCCDET_CLASIF > r_int_ClaAli Then
            If CDbl(Format((r_rst_DatRcc!DEUSOL + r_rst_DatRcc!DEUDOL) / r_dbl_TotDeu * 100, "##0.00")) >= 20 And r_rst_DatRcc!DEUSOL + r_rst_DatRcc!DEUDOL > 100 Then
               r_int_ClaAli = r_rst_DatRcc!RCCDET_CLASIF
            End If
         End If
         r_rst_DatRcc.MoveNext
      Loop
      
      'determina codigo de clasificacion para la provision
      If r_int_ClaInt = r_int_ClaAli Then
         p_int_ClaPrv = r_int_ClaInt
      ElseIf r_int_ClaInt > r_int_ClaAli Then
         p_int_ClaPrv = r_int_ClaInt
      Else
         If r_int_ClaAli = 0 Or r_int_ClaAli = 1 Or r_int_ClaAli = 2 Then
            p_int_ClaPrv = r_int_ClaInt
         Else
            If r_int_ClaInt = 0 Then
               p_int_ClaPrv = r_int_ClaAli - 1
            Else
               p_int_ClaPrv = r_int_ClaAli
            End If
         End If
      End If
   End If
   
   'determina descripcion de la clasificacion para la provision
   If p_Refina = 1 Then
      Select Case p_int_ClaPrv
         Case 0: p_str_ClaPrv = "NOR-REF"
         Case 1: p_str_ClaPrv = "CPP-REF"
         Case 2: p_str_ClaPrv = "DEF-REF"
         Case 3: p_str_ClaPrv = "DUD-REF"
         Case 4: p_str_ClaPrv = "PER-REF"
      End Select
   Else
      If r_int_ClaInt = r_int_ClaAli Then
         Select Case p_int_ClaPrv
            Case 0: p_str_ClaPrv = "NOR"
            Case 1: p_str_ClaPrv = "CPP"
            Case 2: p_str_ClaPrv = "DEF"
            Case 3: p_str_ClaPrv = "DUD"
            Case 4: p_str_ClaPrv = "PER"
         End Select
      Else
         Select Case p_int_ClaPrv
            Case 0: p_str_ClaPrv = "NOR"
            Case 1: p_str_ClaPrv = "CPP"
            Case 2: p_str_ClaPrv = "DEF-ALI"
            Case 3: p_str_ClaPrv = "DUD-ALI"
            Case 4: p_str_ClaPrv = "PER-ALI"
         End Select
      End If
   End If
   
   'Obtiene tasa de provision
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT TIPPRV_PORCEN "
   r_str_Cadena = r_str_Cadena & "  FROM CTB_TIPPRV "
   r_str_Cadena = r_str_Cadena & " WHERE TIPPRV_TIPPRV = '2' "
   r_str_Cadena = r_str_Cadena & "   AND TIPPRV_CLACRE = '13' "
   r_str_Cadena = r_str_Cadena & "   AND TIPPRV_CLFCRE = " & p_int_ClaPrv & " "
   If p_TipGar = 1 Then
      r_str_Cadena = r_str_Cadena & "   AND TIPPRV_CLAGAR = 2 "
   Else
      r_str_Cadena = r_str_Cadena & "   AND TIPPRV_CLAGAR = 1 "
   End If
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_PrcPrv, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_PrcPrv.BOF And r_rst_PrcPrv.EOF) Then
      r_rst_PrcPrv.MoveFirst
      p_TasPrv = r_rst_PrcPrv!TipPrv_Porcen
   End If
   
   r_rst_PrcPrv.Close
   Set r_rst_PrcPrv = Nothing
   
   r_rst_DatRcc.Close
   Set r_rst_DatRcc = Nothing
End Sub

Private Sub fs_GenExc_Actual()
Dim r_obj_Excel     As Excel.Application
Dim r_int_NroFil    As Integer
Dim r_int_Nindex    As Integer
Dim r_dbl_SumCap    As Double
Dim r_dbl_SumPrv    As Double
Dim r_int_nroaux    As Integer
Dim r_bol_FlagOp    As Boolean
    
    r_dbl_SumCap = 0
    r_dbl_SumPrv = 0
    r_int_Nindex = 1
    r_bol_FlagOp = False
    g_str_Parame = gf_QuerySA()
      
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
        Screen.MousePointer = 0
        Exit Sub
    End If
   
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
        g_rst_Princi.Close
        Set g_rst_Princi = Nothing
        MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
        Screen.MousePointer = 0
        Exit Sub
    End If

    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add
    r_int_NroFil = 2
    
    With r_obj_Excel.ActiveSheet
        .Range("A" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
        .Range("A" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Bold = True
        .Range("A" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
        .Range("A" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Size = 14
        .Range("A" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Underline = xlUnderlineStyleSingle
        .Cells(r_int_NroFil, 1) = "PROVISIÓN DE CLIENTES MOROSOS Y ALINEADOS AL " & Day(date) & " DE " & UCase(MonthName(Month(date))) & " DEL " & Year(date)
         
        r_int_NroFil = r_int_NroFil + 2
         
        .Cells(r_int_NroFil, 1) = "PRODUCTO"
        .Cells(r_int_NroFil, 2) = "NOMBRE DEL CLIENTE"
        .Cells(r_int_NroFil, 3) = "DÍAS ATRASO"
        .Cells(r_int_NroFil, 4) = "GARANTÍA"
        .Cells(r_int_NroFil, 5) = "VAL. GARANTÍA"
        .Cells(r_int_NroFil, 6) = "CAPITAL"
        .Cells(r_int_NroFil, 7) = "TASA %"
        .Cells(r_int_NroFil, 8) = "PROVISIÓN"
        .Cells(r_int_NroFil, 9) = "ACTUAL"
        .Cells(r_int_NroFil, 10) = "ALINEADA DEL ÚLTIMO CIERRE"
                   
        .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 12)).Font.Bold = True
        .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 12)).HorizontalAlignment = xlHAlignCenter
        
        .Cells(r_int_NroFil, 1).RowHeight = 30
        .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 12)).VerticalAlignment = xlCenter
        .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 12)).WrapText = True
        
        .Columns("A").ColumnWidth = 11
        .Columns("B").ColumnWidth = 46
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 8
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 9
        .Columns("J").ColumnWidth = 14
        
        .Columns("A").HorizontalAlignment = xlHAlignCenter
        .Columns("C").HorizontalAlignment = xlHAlignCenter
        .Columns("D").HorizontalAlignment = xlHAlignCenter
        .Columns("G").HorizontalAlignment = xlHAlignCenter
        .Columns("I").HorizontalAlignment = xlHAlignCenter
        .Columns("J").HorizontalAlignment = xlHAlignCenter
          
        .Columns("E").Select
        r_obj_Excel.Selection.NumberFormat = "###,###,##0.00"
        .Columns("F").Select
        r_obj_Excel.Selection.NumberFormat = "###,###,##0.00"
        .Columns("H").Select
        r_obj_Excel.Selection.NumberFormat = "###,###,##0.00"
        
        .Cells(1, 1).Select
    End With
   
    Do While r_int_Nindex <= 4
        r_int_nroaux = 1
      
        Do While r_int_nroaux < 3
            r_dbl_SumCap = 0
            r_dbl_SumPrv = 0
            g_rst_Princi.Filter = adFilterNone
            g_rst_Princi.MoveFirst
            
            r_int_NroFil = r_int_NroFil + 1
            
            If r_int_nroaux = 1 Then
                g_rst_Princi.Filter = "CLAACTUAL = " & r_int_Nindex & " AND HIPMAE_TIPGAR > 2"
            Else
                g_rst_Princi.Filter = "CLAACTUAL = " & r_int_Nindex & " AND HIPMAE_TIPGAR < 3"
            End If
            
            If g_rst_Princi.EOF Then
                r_bol_FlagOp = True
            End If
            
            Do While Not g_rst_Princi.EOF
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 1) = Trim(g_rst_Princi!PRODUCTO)
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 2) = Trim(g_rst_Princi!NOMBRECLIENTE)
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 3) = g_rst_Princi!DIASATRASO
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 4) = Trim(g_rst_Princi!TIPOGARAN)
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 5) = CDbl(Format(g_rst_Princi!VALORGARANTIA, "###,###,##0.00"))
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 6) = CDbl(Format(g_rst_Princi!CAPITAL, "###,###,##0.00"))
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 7) = CDbl(Format(g_rst_Princi!TASA, "###,###,##0.00"))
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 8) = CDbl(Format(g_rst_Princi!PROVISION, "###,###,##0.00"))
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 9) = Trim(g_rst_Princi!CLASIFICACTUAL)
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 10) = Trim(g_rst_Princi!ALINEAMULTMES)
                
                r_dbl_SumCap = r_dbl_SumCap + CDbl(Format(g_rst_Princi!CAPITAL, "###,###,##0.00"))
                r_dbl_SumPrv = r_dbl_SumPrv + CDbl(Format(g_rst_Princi!PROVISION, "###,###,##0.00"))
                  
                r_int_NroFil = r_int_NroFil + 1
                g_rst_Princi.MoveNext
                DoEvents
            Loop
            
            If r_dbl_SumCap = 0 Then
                r_int_NroFil = r_int_NroFil - 1
            Else
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 6) = CDbl(Format(r_dbl_SumCap, "###,###,##0.00"))
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 8) = CDbl(Format(r_dbl_SumPrv, "###,###,##0.00"))
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 6).Font.Bold = True
                r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, 8).Font.Bold = True
            End If
            r_int_nroaux = r_int_nroaux + 1
        Loop
        
        If Not r_bol_FlagOp Then r_int_NroFil = r_int_NroFil + 2
        
        r_bol_FlagOp = False
        r_int_Nindex = r_int_Nindex + 1
    Loop

    g_rst_Princi.Close
    Set g_rst_Princi = Nothing

    Screen.MousePointer = 0
    r_obj_Excel.Visible = True
    Set r_obj_Excel = Nothing
End Sub

Private Function gf_NvoCreCom(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer) As Double
Dim r_dbl_MtoAct     As Double

   gf_NvoCreCom = 0
   r_dbl_MtoAct = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUM(DECODE(COMCIE_TIPMON, 1, (COMCIE_PRVGEN+COMCIE_PRVCIC), (COMCIE_PRVGEN+COMCIE_PRVCIC)*COMCIE_TIPCAM)) AS MONTO_TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_COMCIE  "
   g_str_Parame = g_str_Parame & " WHERE COMCIE_PERMES = " & p_MesAct & " "
   g_str_Parame = g_str_Parame & "   AND COMCIE_PERANO = " & p_AnoAct & " "
   g_str_Parame = g_str_Parame & "   AND COMCIE_FECDES >= " & Format(p_AnoAct, "0000") & Format(p_MesAct, "00") & "01 "
   g_str_Parame = g_str_Parame & "   AND COMCIE_CLAPRV = 0  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!MONTO_TOTAL) Then
         r_dbl_MtoAct = CDbl(g_rst_GenAux!MONTO_TOTAL)
      End If
   End If
   
   gf_NvoCreCom = r_dbl_MtoAct
End Function

Private Function gf_AmoCreCom(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer) As Double
Dim r_dbl_MtoAct     As Double
Dim r_dbl_MtoAnt     As Double
Dim r_dbl_MtoDif     As Double
Dim r_dbl_TipCam     As Double
Dim r_rst_PerAnt     As ADODB.Recordset
Dim r_rst_PerAct     As ADODB.Recordset

   gf_AmoCreCom = 0
   r_dbl_MtoAct = 0
   r_dbl_MtoAnt = 0
   r_dbl_MtoDif = 0
   r_dbl_TipCam = 0
   
   'Operaciones del periodo anterior
   If CLng(p_AnoAct & Format(p_MesAct, "00")) < 201903 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT COMCIE_NUMOPE AS NUM_OPERACION, COMCIE_TOTPRE AS IMPORTE"
      g_str_Parame = g_str_Parame & "  FROM CRE_COMCIE "
      g_str_Parame = g_str_Parame & " WHERE COMCIE_PERMES = " & p_MesAnt & " "
      g_str_Parame = g_str_Parame & "   AND COMCIE_PERANO = " & p_AnoAnt & " "
      g_str_Parame = g_str_Parame & "   AND COMCIE_TIPGAR = 1 "
    Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT TRIM(A.CAFCIE_NUMREF) AS NUM_OPERACION, A.CAFCIE_FIAIMP AS IMPORTE, B.GARCIE_TIPGAR "
      g_str_Parame = g_str_Parame & "   FROM TPR_CAFCIE A "
      g_str_Parame = g_str_Parame & "  INNER JOIN TPR_GARCIE B ON B.GARCIE_NUMREF = A.CAFCIE_NUMREF AND A.CAFCIE_PERMES = B.GARCIE_PERMES AND A.CAFCIE_PERANO = B.GARCIE_PERANO "
      g_str_Parame = g_str_Parame & "  WHERE A.CAFCIE_CODPRD = '008' "
      g_str_Parame = g_str_Parame & "    AND A.CAFCIE_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "    AND A.CAFCIE_PERMES = " & p_MesAnt & " "
      g_str_Parame = g_str_Parame & "    AND A.CAFCIE_PERANO = " & p_AnoAnt & " "
      g_str_Parame = g_str_Parame & "    AND B.GARCIE_TIPGAR = 2 "
   End If
   
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_PerAnt, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_PerAnt.BOF And r_rst_PerAnt.EOF) Then
      r_rst_PerAnt.MoveFirst
      Do While Not r_rst_PerAnt.EOF
         r_dbl_MtoAnt = r_rst_PerAnt!IMPORTE
         
         'Busca operaciones del periodo anterior en el periodo actual
         If CLng(p_AnoAct & Format(p_MesAct, "00")) < 201903 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "SELECT COMCIE_NUMOPE AS NUM_OPERACION, COMCIE_TOTPRE AS IMPORTE"
            g_str_Parame = g_str_Parame & "  FROM CRE_COMCIE "
            g_str_Parame = g_str_Parame & " WHERE COMCIE_PERMES = " & p_MesAct & " "
            g_str_Parame = g_str_Parame & "   AND COMCIE_PERANO = " & p_AnoAct & " "
            g_str_Parame = g_str_Parame & "   AND COMCIE_NUMOPE = " & Trim(r_rst_PerAnt!NUM_OPERACION) & " "
         Else
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & " SELECT A.CAFCIE_NUMREF AS NUM_OPERACION, A.CAFCIE_FIAIMP AS IMPORTE, B.GARCIE_TIPGAR "
            g_str_Parame = g_str_Parame & "   FROM TPR_CAFCIE A "
            g_str_Parame = g_str_Parame & "  INNER JOIN TPR_GARCIE B ON TRIM(B.GARCIE_NUMREF) = TRIM(A.CAFCIE_NUMREF) AND A.CAFCIE_PERMES = B.GARCIE_PERMES AND A.CAFCIE_PERANO = B.GARCIE_PERANO "
            g_str_Parame = g_str_Parame & "  WHERE A.CAFCIE_CODPRD = '008' "
            g_str_Parame = g_str_Parame & "    AND A.CAFCIE_SITUAC = 1 "
            g_str_Parame = g_str_Parame & "    AND A.CAFCIE_PERMES = " & p_MesAct & " "
            g_str_Parame = g_str_Parame & "    AND A.CAFCIE_PERANO = " & p_AnoAct & " "
            g_str_Parame = g_str_Parame & "    AND B.GARCIE_TIPGAR = 2 "
            g_str_Parame = g_str_Parame & "    AND TRIM(A.CAFCIE_NUMREF) = " & Trim(r_rst_PerAnt!NUM_OPERACION) & " "
         End If
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_PerAct, 3) Then
            Exit Function
         End If
         
         If Not (r_rst_PerAct.BOF And r_rst_PerAct.EOF) Then
            r_rst_PerAct.MoveFirst
            r_dbl_MtoAct = r_rst_PerAct!IMPORTE
            r_dbl_MtoDif = r_dbl_MtoDif + (r_dbl_MtoAct - r_dbl_MtoAnt)
         Else
            r_dbl_MtoDif = r_dbl_MtoDif - r_dbl_MtoAnt
         End If
         
         r_rst_PerAnt.MoveNext
      Loop
      
      gf_AmoCreCom = r_dbl_MtoDif * (1.3 / 100)
   End If
   
End Function

Private Function gf_DesAmoCreCom(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer, ByRef p_ImpDes As Double, ByRef p_ImpAmo As Double)
   p_ImpDes = 0
   p_ImpAmo = 0

   If CLng(p_AnoAct & Format(p_MesAct, "00")) < 201903 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "  SELECT DISTINCT X.NUM_OPERACION, C.PRVGEN_ACT, B.PRVGEN_ANT, (NVL(C.PRVGEN_ACT,0) - NVL(B.PRVGEN_ANT,0)) AS DIFERENCIA, "
      g_str_Parame = g_str_Parame & "         CASE WHEN NVL(C.PRVGEN_ACT,0) > 0 AND NVL(B.PRVGEN_ANT,0) = 0 THEN 'D' "
      g_str_Parame = g_str_Parame & "         ELSE 'A' END AS TIPO "
      g_str_Parame = g_str_Parame & "    FROM ( SELECT TRIM(A.COMCIE_NUMOPE) AS NUM_OPERACION "
      g_str_Parame = g_str_Parame & "             FROM CRE_COMCIE A "
      g_str_Parame = g_str_Parame & "            WHERE A.COMCIE_PERMES = " & p_MesAct & " "
      g_str_Parame = g_str_Parame & "              AND A.COMCIE_PERANO = " & p_AnoAct & " "
      g_str_Parame = g_str_Parame & "            UNION ALL"
      g_str_Parame = g_str_Parame & "           SELECT TRIM(A.COMCIE_NUMOPE) AS NUM_OPERACION "
      g_str_Parame = g_str_Parame & "             FROM CRE_COMCIE A "
      g_str_Parame = g_str_Parame & "            WHERE A.COMCIE_PERMES = " & p_MesAnt & " "
      g_str_Parame = g_str_Parame & "              AND A.COMCIE_PERANO = " & p_AnoAnt & " "
      g_str_Parame = g_str_Parame & "         ) X "
      g_str_Parame = g_str_Parame & "           LEFT JOIN ( SELECT TRIM(B.COMCIE_NUMOPE) AS NUM_OPERACION, B.COMCIE_PRVGEN AS PRVGEN_ANT "
      g_str_Parame = g_str_Parame & "                         FROM CRE_COMCIE B "
      g_str_Parame = g_str_Parame & "                        WHERE B.COMCIE_PERMES = " & p_MesAnt & " "
      g_str_Parame = g_str_Parame & "                          AND B.COMCIE_PERANO = " & p_AnoAnt & " "
      g_str_Parame = g_str_Parame & "                 ) B ON B.NUM_OPERACION = X.NUM_OPERACION "
      g_str_Parame = g_str_Parame & "           LEFT JOIN ( SELECT TRIM(C.COMCIE_NUMOPE) AS NUM_OPERACION, C.COMCIE_PRVGEN AS PRVGEN_ACT"
      g_str_Parame = g_str_Parame & "                         FROM CRE_COMCIE C "
      g_str_Parame = g_str_Parame & "                        WHERE C.COMCIE_PERMES = " & p_MesAct & " "
      g_str_Parame = g_str_Parame & "                          AND C.COMCIE_PERANO = " & p_AnoAct & " "
      g_str_Parame = g_str_Parame & "                 ) C ON C.NUM_OPERACION = X.NUM_OPERACION "
    Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "  SELECT DISTINCT X.NUM_OPERACION, C.PRVGEN_ACT, B.PRVGEN_ANT, (NVL(C.PRVGEN_ACT,0) - NVL(B.PRVGEN_ANT,0)) AS DIFERENCIA, "
      g_str_Parame = g_str_Parame & "         CASE WHEN NVL(C.PRVGEN_ACT,0) > 0 AND NVL(B.PRVGEN_ANT,0) = 0 THEN 'D' "
      g_str_Parame = g_str_Parame & "         ELSE 'A' END AS TIPO "
      g_str_Parame = g_str_Parame & "    FROM ( SELECT TRIM(A.CAFCIE_NUMREF) AS NUM_OPERACION "
      g_str_Parame = g_str_Parame & "             FROM TPR_CAFCIE A "
      g_str_Parame = g_str_Parame & "            WHERE A.CAFCIE_CODPRD = '008' "
      g_str_Parame = g_str_Parame & "              AND A.CAFCIE_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "              AND A.CAFCIE_PERMES = " & p_MesAct & " "
      g_str_Parame = g_str_Parame & "              AND A.CAFCIE_PERANO = " & p_AnoAct & " "
      g_str_Parame = g_str_Parame & "            UNION ALL"
      g_str_Parame = g_str_Parame & "           SELECT TRIM(A.CAFCIE_NUMREF) AS NUM_OPERACION "
      g_str_Parame = g_str_Parame & "             FROM TPR_CAFCIE A "
      g_str_Parame = g_str_Parame & "            WHERE A.CAFCIE_CODPRD = '008' "
      g_str_Parame = g_str_Parame & "              AND A.CAFCIE_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "              AND A.CAFCIE_PERMES = " & p_MesAnt & " "
      g_str_Parame = g_str_Parame & "              AND A.CAFCIE_PERANO = " & p_AnoAnt & " "
      g_str_Parame = g_str_Parame & "         ) X "
      g_str_Parame = g_str_Parame & "           LEFT JOIN ( SELECT TRIM(B.CAFCIE_NUMREF) AS NUM_OPERACION, B.CAFCIE_PRVGEN AS PRVGEN_ANT "
      g_str_Parame = g_str_Parame & "                         FROM TPR_CAFCIE B "
      g_str_Parame = g_str_Parame & "                        WHERE B.CAFCIE_CODPRD = '008' "
      g_str_Parame = g_str_Parame & "                          AND B.CAFCIE_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "                          AND B.CAFCIE_PERMES = " & p_MesAnt & " "
      g_str_Parame = g_str_Parame & "                          AND B.CAFCIE_PERANO = " & p_AnoAnt & " "
      g_str_Parame = g_str_Parame & "                 ) B ON B.NUM_OPERACION = X.NUM_OPERACION "
      g_str_Parame = g_str_Parame & "           LEFT JOIN ( SELECT TRIM(C.CAFCIE_NUMREF) AS NUM_OPERACION, C.CAFCIE_PRVGEN AS PRVGEN_ACT"
      g_str_Parame = g_str_Parame & "                         FROM TPR_CAFCIE C"
      g_str_Parame = g_str_Parame & "                        WHERE C.CAFCIE_CODPRD = '008'"
      g_str_Parame = g_str_Parame & "                          AND C.CAFCIE_SITUAC = 1"
      g_str_Parame = g_str_Parame & "                          AND C.CAFCIE_PERMES = " & p_MesAct & " "
      g_str_Parame = g_str_Parame & "                          AND C.CAFCIE_PERANO = " & p_AnoAct & " "
      g_str_Parame = g_str_Parame & "                 ) C ON C.NUM_OPERACION = X.NUM_OPERACION "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      
      Do While Not g_rst_GenAux.EOF
         If g_rst_GenAux!TIPO = "D" Then
            p_ImpDes = p_ImpDes + g_rst_GenAux!DIFERENCIA
         Else
            p_ImpAmo = p_ImpAmo + g_rst_GenAux!DIFERENCIA
         End If
         
         g_rst_GenAux.MoveNext
      Loop
   End If
End Function

Private Function gf_QNvoDH(ByVal p_Mes As Integer, ByVal p_Ano As Integer) As Double
   gf_QNvoDH = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ROUND((SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVIG-HIPCIE_INTDIF, (HIPCIE_CAPVIG-HIPCIE_INTDIF)*HIPCIE_TIPCAM))*0.7)/100,2) AS TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERANO = " & p_Ano & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERMES = " & p_Mes & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_FECDES >='" & p_Ano & Format(p_Mes, "00") & "01' "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_FECDES <='" & p_Ano & Format(p_Mes, "00") & "31' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!total) Then
         gf_QNvoDH = CDbl(g_rst_GenAux!total)
      End If
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Function

Private Function gf_PorHip(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer, ByVal p_TipCal As Integer) As Double
Dim r_dbl_MtoAct     As Double
Dim r_dbl_MtoAnt     As Double
Dim r_dbl_MtoTrf     As Double
Dim r_dbl_TipCam     As Double

   gf_PorHip = 0
   r_dbl_MtoAct = 0
   r_dbl_MtoAnt = 0
   r_dbl_TipCam = 0
   r_dbl_MtoTrf = 0
   
   'Monto del periodo actual
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC), (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC)*HIPCIE_TIPCAM)),2) AS MONTO_TOTAL, "
   g_str_Parame = g_str_Parame & "       MAX(HIPCIE_TIPCAM) AS TIPO_CAMBIO"
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE  "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & p_MesAct & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & p_AnoAct & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_FECDES < " & Format(p_AnoAct, "0000") & Format(p_MesAct, "00") & "01 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!MONTO_TOTAL) Then
         r_dbl_MtoAct = CDbl(g_rst_GenAux!MONTO_TOTAL)
         r_dbl_TipCam = CDbl(g_rst_GenAux!TIPO_CAMBIO)
      End If
   End If
   
   'Monto de los transferidos (si hubieran)
   g_str_Parame = ""
   'g_str_Parame = g_str_Parame & "SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC+HIPCIE_PRVVOL), (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC+HIPCIE_PRVVOL)*" & CStr(r_dbl_TipCam) & ")),2) AS MONTO_TOTAL "
   g_str_Parame = g_str_Parame & "SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC), (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC)*" & CStr(r_dbl_TipCam) & ")),2) AS MONTO_TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE  "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & p_MesAnt & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & p_AnoAnt & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_NUMOPE IN (SELECT HIPMAE_NUMOPE FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & "                          WHERE HIPMAE_SITUAC = 6 "
   g_str_Parame = g_str_Parame & "                            AND HIPMAE_FECCAN >= " & Format(p_AnoAct, "0000") & Format(p_MesAct, "00") & "01"
   g_str_Parame = g_str_Parame & "                            AND HIPMAE_FECCAN <= " & Format(p_AnoAct, "0000") & Format(p_MesAct, "00") & "31)"
   'g_str_Parame = g_str_Parame & "                            AND HIPMAE_FECCAN <= " & Format(IIf(p_MesAct = 12, p_AnoAct + 1, p_AnoAct), "0000") & Format(IIf(p_MesAct = 12, 1, p_MesAct + 1), "00") & "01)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!MONTO_TOTAL) Then
         r_dbl_MtoTrf = CDbl(g_rst_GenAux!MONTO_TOTAL)
      End If
   End If
      
   'Monto del periodo anterior
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC), (HIPCIE_PRVGEN+HIPCIE_PRVGEN_RC+HIPCIE_PRVCIC+HIPCIE_PRVCIC_RC)*" & CStr(r_dbl_TipCam) & ")),2) AS MONTO_TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE  "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & p_MesAnt & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & p_AnoAnt & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!MONTO_TOTAL) Then
         r_dbl_MtoAnt = CDbl(g_rst_GenAux!MONTO_TOTAL)
      End If
   End If
   
   If p_TipCal = 1 Then
      gf_PorHip = Format(0 - r_dbl_MtoTrf, "###,##0.00")
   Else
      gf_PorHip = Format((r_dbl_MtoAct + r_dbl_MtoTrf) - r_dbl_MtoAnt, "###,##0.00")
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Function

Private Function gf_SalFav(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer) As Double
Dim r_dbl_TipCam     As Double

   gf_SalFav = 0
   
   'Obtiene el tipo de cambio del mes
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DISTINCT HIPCIE_TIPCAM "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERANO = " & p_AnoAct & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERMES = " & p_MesAct & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      r_dbl_TipCam = CDbl(g_rst_GenAux!HIPCIE_TIPCAM)
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ROUND( SUM( DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*" & CStr(r_dbl_TipCam) & ") )* -1 ,2) AS TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERANO = " & p_AnoAnt & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERMES = " & p_MesAnt & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_CLAPRV <> 0 "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_NUMOPE NOT IN ( "
   g_str_Parame = g_str_Parame & "          SELECT HIPCIE_NUMOPE "
   g_str_Parame = g_str_Parame & "            FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & "           WHERE HIPCIE_PERANO = " & p_AnoAct & " "
   g_str_Parame = g_str_Parame & "             AND HIPCIE_PERMES = " & p_MesAct & " "
   g_str_Parame = g_str_Parame & "             AND HIPCIE_CLAPRV <> 0 )"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
    
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!total) Then
         gf_SalFav = CDbl(g_rst_GenAux!total)
      End If
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Function

Private Function gf_PrvVol(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer) As Double
Dim r_dbl_MtoAct     As Double
Dim r_dbl_MtoAnt     As Double
Dim r_dbl_MtoDif     As Double
Dim r_dbl_TipCam     As Double
Dim r_rst_PerAnt     As ADODB.Recordset
Dim r_rst_PerAct     As ADODB.Recordset

   gf_PrvVol = 0
   r_dbl_MtoAct = 0
   r_dbl_MtoAnt = 0
   r_dbl_TipCam = 0
   
   'Monto del periodo actual
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVVOL, HIPCIE_PRVVOL*HIPCIE_TIPCAM)),2) AS MONTO_TOTAL, "
   g_str_Parame = g_str_Parame & "       MAX(HIPCIE_TIPCAM) AS TIPO_CAMBIO"
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE  "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & p_MesAct & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & p_AnoAct & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!MONTO_TOTAL) Then
         r_dbl_MtoAct = CDbl(g_rst_GenAux!MONTO_TOTAL)
         r_dbl_TipCam = CDbl(g_rst_GenAux!TIPO_CAMBIO)
      End If
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   'Monto del periodo anterior
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_PRVVOL), (HIPCIE_PRVVOL)*" & CStr(r_dbl_TipCam) & ")),2) AS MONTO_TOTAL "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE  "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & p_MesAnt & " "
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & p_AnoAnt & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      If Not IsNull(g_rst_GenAux!MONTO_TOTAL) Then
         r_dbl_MtoAnt = CDbl(g_rst_GenAux!MONTO_TOTAL)
      End If
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   gf_PrvVol = Format(r_dbl_MtoAct - r_dbl_MtoAnt, "###,##0.00")
End Function

''QUERY 1
Private Function gf_Query(ByVal p_MesAct1 As Integer, ByVal p_AnoAct1 As Integer, ByVal p_MesAnt1 As Integer, ByVal p_AnoAnt1 As Integer) As String
    gf_Query = ""
    gf_Query = gf_Query & " SELECT * FROM ( "
    gf_Query = gf_Query & " WITH QUERY1 AS ( "
    gf_Query = gf_Query & "         SELECT HIPCIE_NUMOPE, HIPCIE_TDOCLI, HIPCIE_NDOCLI, HIPCIE_CLACLI, HIPCIE_NUMOPE AS OPERACION, HIPCIE_CLAALI, "
    gf_Query = gf_Query & "                HIPCIE_CLAPRV, HIPCIE_TIPGAR, HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, HIPCIE_CODPRD, HIPCIE_CODSUB, HIPMAE_HIPMTZ, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CODPRD='001' THEN 'CRC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='002' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='003' THEN 'CME' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='004' THEN 'MIH' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='006' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='007' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='009' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='010' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='011' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='012' THEN 'UAN' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='013' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='014' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='015' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='016' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='017' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='018' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='019' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='021' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='022' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='023' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='024' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='025' THEN 'MIV' END AS PRODUCTO, "
    gf_Query = gf_Query & "                TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRECLIENTE, "
    gf_Query = gf_Query & "                HIPCIE_DIAMOR AS DIASATRASO, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_TIPGAR=1 THEN 'HIP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=2 THEN 'BLQ' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=3 THEN 'FS'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=4 THEN 'CF'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=5 THEN 'CP'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=6 THEN 'RF'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=8 THEN 'FSHM'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=9 THEN 'GH' END AS TIPOGARAN, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_TIPGAR IN (1,2,4) THEN DECODE(HIPCIE_MONGAR,1,HIPCIE_MTOGAR,(HIPCIE_MTOGAR)*HIPCIE_TIPCAM) "
    gf_Query = gf_Query & "                     ELSE 0 END AS VALORGARANTIA, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF,(HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF)*HIPCIE_TIPCAM), 2) AS CAPITAL, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON,(HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM), 2) AS CAPITAL_BAL, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_PRVVOL,(HIPCIE_PRVVOL)*HIPCIE_TIPCAM), 2) AS PROV_VOLUNT1, "
    gf_Query = gf_Query & "                NVL((SELECT TIPPRV_PORCEN "
    gf_Query = gf_Query & "                       FROM CTB_TIPPRV "
    gf_Query = gf_Query & "                      WHERE TIPPRV_TIPPRV = '2' "
    gf_Query = gf_Query & "                        AND TIPPRV_CLACRE = '13' "
    gf_Query = gf_Query & "                        AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
    gf_Query = gf_Query & "                        AND TIPPRV_CLAGAR = DECODE(HIPCIE_TIPGAR,1,2, DECODE(HIPCIE_TIPGAR,4,2,1) )), 0) AS TASA, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*HIPCIE_TIPCAM),2) AS PROVISION, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
    gf_Query = gf_Query & "                        END "
    gf_Query = gf_Query & "                     ELSE "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN "
    gf_Query = gf_Query & "                             CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                                  WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                                  WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                                  WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                                  WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
    gf_Query = gf_Query & "                             END "
    gf_Query = gf_Query & "                        ELSE "
    gf_Query = gf_Query & "                             CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                                  WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                                  WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
    gf_Query = gf_Query & "                                  WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
    gf_Query = gf_Query & "                                  WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
    gf_Query = gf_Query & "                             END "
    gf_Query = gf_Query & "                        END "
    gf_Query = gf_Query & "                END AS CLASIFICACION, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
    gf_Query = gf_Query & "                END AS CLAINT1, "
    gf_Query = gf_Query & "                CASE WHEN HIPMAE_HIPMTZ ='1' THEN 'NO APLICA' "
    gf_Query = gf_Query & "                     WHEN HIPMAE_HIPMTZ ='2' THEN 'LEVANTADA'"
    gf_Query = gf_Query & "                     WHEN HIPMAE_HIPMTZ ='3' THEN 'NO LEVANTADA'"
    gf_Query = gf_Query & "                END AS MATRIZ"
    gf_Query = gf_Query & "           FROM CRE_HIPCIE "
    gf_Query = gf_Query & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
    gf_Query = gf_Query & "          INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCIE_NUMOPE "
    gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = " & p_MesAct1
    gf_Query = gf_Query & "            AND HIPCIE_PERANO = " & p_AnoAct1
    gf_Query = gf_Query & "            AND HIPCIE_CLAPRV <> 0 "
    gf_Query = gf_Query & "         UNION "
    gf_Query = gf_Query & "         SELECT HIPCIE_NUMOPE, HIPCIE_TDOCLI, HIPCIE_NDOCLI, HIPCIE_CLACLI, HIPCIE_NUMOPE AS OPERACION, HIPCIE_CLAALI, "
    gf_Query = gf_Query & "                HIPCIE_CLAPRV, HIPCIE_TIPGAR, HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, HIPCIE_CODPRD, HIPCIE_CODSUB, HIPMAE_HIPMTZ, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CODPRD='001' THEN 'CRC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='002' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='003' THEN 'CME' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='004' THEN 'MIH' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='006' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='007' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='009' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='010' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='011' THEN 'MIC' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='012' THEN 'UAN' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='013' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='014' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='015' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='016' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='017' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='018' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='019' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='021' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='022' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='023' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='024' THEN 'MIV' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CODPRD='025' THEN 'MIV' END AS PRODUCTO, "
    gf_Query = gf_Query & "                TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRECLIENTE, "
    gf_Query = gf_Query & "                HIPCIE_DIAMOR AS DIASATRASO, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_TIPGAR=1 THEN 'HIP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=2 THEN 'BLQ' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=3 THEN 'FS'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=4 THEN 'CF'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=5 THEN 'CP'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=6 THEN 'RF'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=8 THEN 'FSHM' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=9 THEN 'GH' END AS TIPOGARAN, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_TIPGAR IN (1,2,4) THEN DECODE(HIPCIE_MONGAR,1,HIPCIE_MTOGAR,(HIPCIE_MTOGAR)*HIPCIE_TIPCAM) "
    gf_Query = gf_Query & "                     ELSE 0 END AS VALORGARANTIA, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF,(HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF)*HIPCIE_TIPCAM), 2) AS CAPITAL, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON,(HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM), 2) AS CAPITAL_BAL, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_PRVVOL,(HIPCIE_PRVVOL)*HIPCIE_TIPCAM), 2) AS PROV_VOLUNT1, "
    gf_Query = gf_Query & "                NVL((SELECT TIPPRV_PORCEN "
    gf_Query = gf_Query & "                       FROM CTB_TIPPRV "
    gf_Query = gf_Query & "                      WHERE TIPPRV_TIPPRV = '2' "
    gf_Query = gf_Query & "                        AND TIPPRV_CLACRE = '13' "
    gf_Query = gf_Query & "                        AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
    gf_Query = gf_Query & "                        AND TIPPRV_CLAGAR = DECODE(HIPCIE_TIPGAR,1,2, DECODE(HIPCIE_TIPGAR,4,2,1) )), 0) AS TASA, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*HIPCIE_TIPCAM),2) AS PROVISION, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
    gf_Query = gf_Query & "                        END "
    gf_Query = gf_Query & "                     ELSE "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN "
    gf_Query = gf_Query & "                           CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                                WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                                WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                                WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                                WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
    gf_Query = gf_Query & "                           END "
    gf_Query = gf_Query & "                        ELSE "
    gf_Query = gf_Query & "                           CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                                WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                                WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
    gf_Query = gf_Query & "                                WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
    gf_Query = gf_Query & "                                WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
    gf_Query = gf_Query & "                           END "
    gf_Query = gf_Query & "                     END "
    gf_Query = gf_Query & "                END AS CLASIFICACION, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
    gf_Query = gf_Query & "                END AS CLAINT1, "
    gf_Query = gf_Query & "                CASE WHEN HIPMAE_HIPMTZ ='1' THEN 'NO APLICA' "
    gf_Query = gf_Query & "                     WHEN HIPMAE_HIPMTZ ='2' THEN 'LEVANTADA'"
    gf_Query = gf_Query & "                     WHEN HIPMAE_HIPMTZ ='3' THEN 'NO LEVANTADA'"
    gf_Query = gf_Query & "                END AS MATRIZ"
    gf_Query = gf_Query & "           FROM CRE_HIPCIE "
    gf_Query = gf_Query & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
    gf_Query = gf_Query & "          INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCIE_NUMOPE "
    gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = " & p_MesAct1
    gf_Query = gf_Query & "            AND HIPCIE_PERANO = " & p_AnoAct1
    gf_Query = gf_Query & "            AND HIPCIE_PRVVOL > 0 "
    gf_Query = gf_Query & "            AND HIPCIE_CLAALI = 0 "
    
    'Adicional 28/09/2015
    gf_Query = gf_Query & "         UNION "
    gf_Query = gf_Query & "   SELECT HIPCIE_NUMOPE, HIPCIE_TDOCLI, HIPCIE_NDOCLI, HIPCIE_CLACLI, HIPCIE_NUMOPE AS OPERACION, HIPCIE_CLAALI, "
    gf_Query = gf_Query & "          HIPCIE_CLAPRV, HIPCIE_TIPGAR, HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, HIPCIE_CODPRD, HIPCIE_CODSUB, HIPMAE_HIPMTZ, "
    gf_Query = gf_Query & "          CASE WHEN HIPCIE_CODPRD='001' THEN 'CRC' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='002' THEN 'MIC' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='003' THEN 'CME' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='004' THEN 'MIH' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='006' THEN 'MIC' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='007' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='009' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='010' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='011' THEN 'MIC' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='012' THEN 'UAN' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='013' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='014' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='015' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='016' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='017' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='018' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='019' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='021' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='022' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='023' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='024' THEN 'MIV' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CODPRD='025' THEN 'MIV' END AS PRODUCTO, "
    gf_Query = gf_Query & "          TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRECLIENTE, "
    gf_Query = gf_Query & "          HIPCIE_DIAMOR AS DIASATRASO, "
    gf_Query = gf_Query & "          CASE WHEN HIPCIE_TIPGAR=1 THEN 'HIP' "
    gf_Query = gf_Query & "               WHEN HIPCIE_TIPGAR=2 THEN 'BLQ' "
    gf_Query = gf_Query & "               WHEN HIPCIE_TIPGAR=3 THEN 'FS' "
    gf_Query = gf_Query & "               WHEN HIPCIE_TIPGAR=4 THEN 'CF' "
    gf_Query = gf_Query & "               WHEN HIPCIE_TIPGAR=5 THEN 'CP' "
    gf_Query = gf_Query & "               WHEN HIPCIE_TIPGAR=6 THEN 'RF' "
    gf_Query = gf_Query & "               WHEN HIPCIE_TIPGAR=8 THEN 'FSHM' "
    gf_Query = gf_Query & "               WHEN HIPCIE_TIPGAR=9 THEN 'GH' END AS TIPOGARAN, "
    gf_Query = gf_Query & "          CASE WHEN HIPCIE_TIPGAR IN (1,2,4) THEN DECODE(HIPCIE_MONGAR,1,HIPCIE_MTOGAR,(HIPCIE_MTOGAR)*HIPCIE_TIPCAM) "
    gf_Query = gf_Query & "               ELSE 0 END AS VALORGARANTIA, "
    gf_Query = gf_Query & "          ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF,(HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF)*HIPCIE_TIPCAM), 2) AS CAPITAL, "
    gf_Query = gf_Query & "          ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON,(HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM), 2) AS CAPITAL_BAL, "
    gf_Query = gf_Query & "          ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_PRVVOL,(HIPCIE_PRVVOL)*HIPCIE_TIPCAM), 2) AS PROV_VOLUNT1, "
    gf_Query = gf_Query & "          NVL((SELECT TIPPRV_PORCEN "
    gf_Query = gf_Query & "                 FROM CTB_TIPPRV "
    gf_Query = gf_Query & "                WHERE TIPPRV_TIPPRV = '2' "
    gf_Query = gf_Query & "                  AND TIPPRV_CLACRE = '13' "
    gf_Query = gf_Query & "                  AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
    gf_Query = gf_Query & "                  AND TIPPRV_CLAGAR = DECODE(HIPCIE_TIPGAR,1,2, DECODE(HIPCIE_TIPGAR,4,2,1) )), 0) AS TASA, "
    gf_Query = gf_Query & "          ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*HIPCIE_TIPCAM),2) AS PROVISION, "
    gf_Query = gf_Query & "          CASE WHEN HIPCIE_FLGREF = 1 THEN "
    gf_Query = gf_Query & "                  CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR-REF' "
    gf_Query = gf_Query & "                       WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
    gf_Query = gf_Query & "                       WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
    gf_Query = gf_Query & "                       WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
    gf_Query = gf_Query & "                       WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
    gf_Query = gf_Query & "                  END "
    gf_Query = gf_Query & "               ELSE "
    gf_Query = gf_Query & "                 CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN "
    gf_Query = gf_Query & "                    CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                         WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                         WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                         WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                         WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
    gf_Query = gf_Query & "                    END "
    gf_Query = gf_Query & "                 ELSE "
    gf_Query = gf_Query & "                    CASE WHEN HIPCIE_CLAPRV = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                         WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                         WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
    gf_Query = gf_Query & "                         WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
    gf_Query = gf_Query & "                         WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
    gf_Query = gf_Query & "                    END "
    gf_Query = gf_Query & "               END "
    gf_Query = gf_Query & "          END AS CLASIFICACION, "
    gf_Query = gf_Query & "          CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
    gf_Query = gf_Query & "               WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
    gf_Query = gf_Query & "          END AS CLAINT1, "
    gf_Query = gf_Query & "          CASE WHEN HIPMAE_HIPMTZ ='1' THEN 'NO APLICA' "
    gf_Query = gf_Query & "               WHEN HIPMAE_HIPMTZ ='2' THEN 'LEVANTADA'"
    gf_Query = gf_Query & "               WHEN HIPMAE_HIPMTZ ='3' THEN 'NO LEVANTADA'"
    gf_Query = gf_Query & "           END AS MATRIZ"
    gf_Query = gf_Query & "     FROM CRE_HIPCIE A "
    gf_Query = gf_Query & "    INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
    gf_Query = gf_Query & "          INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCIE_NUMOPE "
    gf_Query = gf_Query & "    WHERE A.HIPCIE_PERMES = " & p_MesAct1
    gf_Query = gf_Query & "      AND A.HIPCIE_PERANO = " & p_AnoAct1
    gf_Query = gf_Query & "      AND A.HIPCIE_NUMOPE IN (SELECT X.HIPCIE_NUMOPE FROM CRE_HIPCIE X WHERE X.HIPCIE_PERMES = " & p_MesAnt1 & " AND X.HIPCIE_PERANO = " & p_AnoAnt1 & " AND X.HIPCIE_PRVVOL > 0) "
    gf_Query = gf_Query & "      AND A.HIPCIE_PRVVOL = 0 "
    gf_Query = gf_Query & "      AND A.HIPCIE_CLAPRV = 0 "
    gf_Query = gf_Query & "            ), "
    
    gf_Query = gf_Query & " QUERY2 AS ("
    gf_Query = gf_Query & "         SELECT HIPCIE_NUMOPE AS NUMOPE, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesAct1 & " AND HIPCIE_PERANO = " & p_AnoAct1 & ")),2) AS PROVISION2, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVVOL, HIPCIE_PRVVOL*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesAct1 & " AND HIPCIE_PERANO = " & p_AnoAct1 & ")),2) AS PROV_VOLUNT2, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
    gf_Query = gf_Query & "                        END "
    gf_Query = gf_Query & "                     ELSE "
    gf_Query = gf_Query & "                       CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN  "
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
    gf_Query = gf_Query & "                          END"
    gf_Query = gf_Query & "                       ELSE"
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
    gf_Query = gf_Query & "                          END"
    gf_Query = gf_Query & "                     END"
    gf_Query = gf_Query & "                END AS CLASIFICACION2, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
    gf_Query = gf_Query & "                END AS CLAINT2 "
    gf_Query = gf_Query & "           FROM CRE_HIPCIE "
    gf_Query = gf_Query & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
    gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = " & p_MesAnt1
    gf_Query = gf_Query & "            AND HIPCIE_PERANO = " & p_AnoAnt1
    gf_Query = gf_Query & "            AND HIPCIE_CLAPRV <> 0 "
    gf_Query = gf_Query & "         UNION "
    gf_Query = gf_Query & "         SELECT HIPCIE_NUMOPE AS NUMOPE, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesAct1 & " AND HIPCIE_PERANO = " & p_AnoAct1 & ")),2) AS PROVISION2, "
    gf_Query = gf_Query & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVVOL, HIPCIE_PRVVOL*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesAct1 & " AND HIPCIE_PERANO = " & p_AnoAct1 & ")),2) AS PROV_VOLUNT2, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
    gf_Query = gf_Query & "                        END "
    gf_Query = gf_Query & "                     ELSE "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN  "
    gf_Query = gf_Query & "                           CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
    gf_Query = gf_Query & "                          END"
    gf_Query = gf_Query & "                       ELSE"
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
    gf_Query = gf_Query & "                          END"
    gf_Query = gf_Query & "                     END"
    gf_Query = gf_Query & "                END AS CLASIFICACION2, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
    gf_Query = gf_Query & "                END AS CLAINT2 "
    gf_Query = gf_Query & "           FROM CRE_HIPCIE "
    gf_Query = gf_Query & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
    gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = " & p_MesAnt1
    gf_Query = gf_Query & "            AND HIPCIE_PERANO = " & p_AnoAnt1
    gf_Query = gf_Query & "            AND HIPCIE_PRVVOL > 0 "
    gf_Query = gf_Query & "            AND HIPCIE_CLAALI >= 0 )"
    
    gf_Query = gf_Query & "  SELECT PRODUCTO, NOMBRECLIENTE, DIASATRASO, TIPOGARAN "
    gf_Query = gf_Query & "         ,ROUND(VALORGARANTIA,2) AS VALORGARANTIA, TASA "
    gf_Query = gf_Query & "         ,ROUND(CAPITAL,2) AS CAPITAL, ROUND(CAPITAL_BAL,2) AS CAPITAL_BAL "
    gf_Query = gf_Query & "         ,ROUND(PROVISION,2) AS PROVISION, CLASIFICACION, CLAINT1 "
    gf_Query = gf_Query & "         ,ROUND(NVL(PROVISION2,0),2) AS PROVISION2, CLASIFICACION2, CLAINT2 "
    gf_Query = gf_Query & "         ,ROUND((NVL(PROVISION,0) - NVL(PROVISION2,0)),2) AS AJUSTE "
    gf_Query = gf_Query & "         ,HIPCIE_CLAPRV, HIPCIE_TIPGAR, HIPCIE_TDOCLI, HIPCIE_NDOCLI "
    gf_Query = gf_Query & "         ,HIPCIE_CLACLI, HIPCIE_CLAALI, PROV_VOLUNT1, PROV_VOLUNT2 "
    gf_Query = gf_Query & "         ,HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, OPERACION, "
    
    '09/05/2016
    gf_Query = gf_Query & "         (CASE WHEN H.SOLINM_TABPRY IS NOT NULL THEN "
    gf_Query = gf_Query & "               CASE WHEN H.SOLINM_TABPRY = 2 THEN "
    gf_Query = gf_Query & "                    CASE WHEN H.SOLINM_PRYCOD IS NOT NULL THEN "
    gf_Query = gf_Query & "                         CASE WHEN LENGTH (H.SOLINM_PRYCOD) > 0 THEN "
    gf_Query = gf_Query & "                              (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD) "
    gf_Query = gf_Query & "                         ELSE "
    gf_Query = gf_Query & "                              CASE WHEN LENGTH (H.SOLINM_PRYNOM) > 0 THEN TRIM(H.SOLINM_PRYNOM) END "
    gf_Query = gf_Query & "                          END "
    gf_Query = gf_Query & "                    ELSE CASE WHEN LENGTH (H.SOLINM_PRYCOD) > 0 THEN "
    gf_Query = gf_Query & "                              (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD) "
    gf_Query = gf_Query & "                         ELSE"
    gf_Query = gf_Query & "                              CASE WHEN H.SOLINM_PRYNOM IS NOT NULL THEN "
    gf_Query = gf_Query & "                                Trim (H.SOLINM_PRYNOM) "
    gf_Query = gf_Query & "                              ELSE '' "
    gf_Query = gf_Query & "                               END "
    gf_Query = gf_Query & "                          END "
    gf_Query = gf_Query & "                     END "
    gf_Query = gf_Query & "               ELSE CASE WHEN H.SOLINM_PRYCOD IS NOT NULL THEN "
    gf_Query = gf_Query & "                          (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD) "
    gf_Query = gf_Query & "                    ELSE "
    gf_Query = gf_Query & "                          CASE WHEN H.SOLINM_PRYNOM IS NOT NULL THEN "
    gf_Query = gf_Query & "                            Trim (H.SOLINM_PRYNOM) "
    gf_Query = gf_Query & "                          ELSE "
    gf_Query = gf_Query & "                            '' "
    gf_Query = gf_Query & "                           END "
    gf_Query = gf_Query & "                    END "
    gf_Query = gf_Query & "                END "
    gf_Query = gf_Query & "          ELSE CASE WHEN H.SOLINM_PRYCOD IS NOT NULL THEN "
    gf_Query = gf_Query & "                  (SELECT DATGEN_TITULO FROM PRY_DATGEN WHERE DATGEN_CODIGO = H.SOLINM_PRYCOD) "
    gf_Query = gf_Query & "                ELSE "
    gf_Query = gf_Query & "                  '' "
    gf_Query = gf_Query & "                END "
    gf_Query = gf_Query & "           END) AS NOMBRE_PROYECTO, "
    gf_Query = gf_Query & "           CASE WHEN INSTR(TRIM(I.SUBPRD_DESCRI),'MICROEMPRESARIO') > 0 THEN 'X' END  AS MICROEMPRESARIO, MATRIZ "

    gf_Query = gf_Query & "   FROM QUERY1 "
    gf_Query = gf_Query & "         LEFT JOIN QUERY2 ON (HIPCIE_NUMOPE = NUMOPE) "
    
    '09/05/2016
    gf_Query = gf_Query & "         INNER JOIN CRE_SUBPRD I ON I.SUBPRD_CODPRD = HIPCIE_CODPRD AND I.SUBPRD_CODSUB = HIPCIE_CODSUB"
    gf_Query = gf_Query & "         LEFT JOIN CRE_HIPMAE C ON C.HIPMAE_NUMOPE = OPERACION"
    gf_Query = gf_Query & "         LEFT JOIN CRE_SOLINM H ON H.SOLINM_NUMSOL = C.HIPMAE_NUMSOL"
 
    gf_Query = gf_Query & "  ORDER BY HIPCIE_CLAPRV, TIPOGARAN, DIASATRASO "
    gf_Query = gf_Query & "  )"
End Function

'Private Function gf_Query_Old(ByVal p_MesAct1 As Integer, ByVal p_AnoAct1 As Integer, ByVal p_MesAnt1 As Integer, ByVal p_AnoAnt1 As Integer) As String
'    gf_Query_Old = ""
'    gf_Query_Old = gf_Query_Old & " SELECT * FROM ( "
'    gf_Query_Old = gf_Query_Old & " WITH QUERY1 AS ( "
'    gf_Query_Old = gf_Query_Old & "         SELECT HIPCIE_NUMOPE, HIPCIE_TDOCLI, HIPCIE_NDOCLI, HIPCIE_CLACLI, HIPCIE_NUMOPE AS OPERACION, HIPCIE_CLAALI, "
'    gf_Query_Old = gf_Query_Old & "                HIPCIE_CLAPRV, HIPCIE_TIPGAR, HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, "
'    gf_Query_Old = gf_Query_Old & "                CASE WHEN HIPCIE_CODPRD='001' THEN 'CRC' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='002' THEN 'MIC' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='003' THEN 'CME' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='004' THEN 'MIH' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='006' THEN 'MIC' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='007' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='009' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='010' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='011' THEN 'MIC' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='012' THEN 'UAN' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='013' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='014' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='015' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='016' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='017' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='018' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='019' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='021' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='022' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='023' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='024' THEN 'MIV' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CODPRD='025' THEN 'MIV' END AS PRODUCTO, "
'    gf_Query_Old = gf_Query_Old & "                TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRECLIENTE, "
'    gf_Query_Old = gf_Query_Old & "                HIPCIE_DIAMOR AS DIASATRASO, "
'    gf_Query_Old = gf_Query_Old & "                CASE WHEN HIPCIE_TIPGAR=1 THEN 'HIP' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_TIPGAR=2 THEN 'HIP' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_TIPGAR=3 THEN 'FS'  "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_TIPGAR=4 THEN 'CF'  "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_TIPGAR=5 THEN 'CP'  "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_TIPGAR=6 THEN 'RF'  "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_TIPGAR=8 THEN 'FSHM' END AS TIPOGARAN, "
'    gf_Query_Old = gf_Query_Old & "                CASE WHEN HIPCIE_TIPGAR IN (1,2) THEN DECODE(HIPCIE_MONGAR,1,HIPCIE_MTOGAR,(HIPCIE_MTOGAR)*HIPCIE_TIPCAM) "
'    gf_Query_Old = gf_Query_Old & "                     ELSE 0 END AS VALORGARANTIA, "
'    gf_Query_Old = gf_Query_Old & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF,(HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF)*HIPCIE_TIPCAM), 2) AS CAPITAL, "
'    gf_Query_Old = gf_Query_Old & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON,(HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM), 2) AS CAPITAL_BAL, "
'    gf_Query_Old = gf_Query_Old & "                ROUND(DECODE(HIPCIE_TIPMON,1,HIPCIE_PRVVOL,(HIPCIE_PRVVOL)*HIPCIE_TIPCAM), 2) AS PROV_VOLUNT1, "
'    gf_Query_Old = gf_Query_Old & "                (SELECT TIPPRV_PORCEN "
'    gf_Query_Old = gf_Query_Old & "                   FROM CTB_TIPPRV "
'    gf_Query_Old = gf_Query_Old & "                  WHERE TIPPRV_TIPPRV = '2' "
'    gf_Query_Old = gf_Query_Old & "                    AND TIPPRV_CLACRE = '13' "
'    gf_Query_Old = gf_Query_Old & "                    AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
'    gf_Query_Old = gf_Query_Old & "                    AND TIPPRV_CLAGAR = DECODE(HIPCIE_TIPGAR,1,2,1)) AS TASA, "
'    gf_Query_Old = gf_Query_Old & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*HIPCIE_TIPCAM),2) AS PROVISION, "
'    gf_Query_Old = gf_Query_Old & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
'    gf_Query_Old = gf_Query_Old & "                        CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
'    gf_Query_Old = gf_Query_Old & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
'    gf_Query_Old = gf_Query_Old & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
'    gf_Query_Old = gf_Query_Old & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
'    gf_Query_Old = gf_Query_Old & "                        END "
'    gf_Query_Old = gf_Query_Old & "                     ELSE "
'    gf_Query_Old = gf_Query_Old & "                       CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN "
'    gf_Query_Old = gf_Query_Old & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
'    gf_Query_Old = gf_Query_Old & "                          END "
'    gf_Query_Old = gf_Query_Old & "                       ELSE "
'    gf_Query_Old = gf_Query_Old & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
'    gf_Query_Old = gf_Query_Old & "                          END "
'    gf_Query_Old = gf_Query_Old & "                     END "
'    gf_Query_Old = gf_Query_Old & "                END AS CLASIFICACION, "
'    gf_Query_Old = gf_Query_Old & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
'    gf_Query_Old = gf_Query_Old & "                END AS CLAINT1 "
'    gf_Query_Old = gf_Query_Old & "           FROM CRE_HIPCIE "
'    gf_Query_Old = gf_Query_Old & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
'    gf_Query_Old = gf_Query_Old & "          WHERE HIPCIE_PERMES = " & p_MesAct1
'    gf_Query_Old = gf_Query_Old & "            AND HIPCIE_PERANO = " & p_AnoAct1
'    gf_Query_Old = gf_Query_Old & "            AND HIPCIE_CLAPRV <> 0"
'    gf_Query_Old = gf_Query_Old & "          ORDER BY HIPCIE_CLAPRV, TIPOGARAN, DIASATRASO), "
'    gf_Query_Old = gf_Query_Old & " QUERY2 AS ("
'    gf_Query_Old = gf_Query_Old & "         SELECT HIPCIE_NUMOPE AS NUMOPE, "
'    gf_Query_Old = gf_Query_Old & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesAct1 & " AND HIPCIE_PERANO = " & p_AnoAct1 & ")),2) AS PROVISION2, "
'    gf_Query_Old = gf_Query_Old & "                ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVVOL, HIPCIE_PRVVOL*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesAct1 & " AND HIPCIE_PERANO = " & p_AnoAct1 & ")),2) AS PROV_VOLUNT2, "
'    gf_Query_Old = gf_Query_Old & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
'    gf_Query_Old = gf_Query_Old & "                        CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
'    gf_Query_Old = gf_Query_Old & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
'    gf_Query_Old = gf_Query_Old & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
'    gf_Query_Old = gf_Query_Old & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
'    gf_Query_Old = gf_Query_Old & "                        END "
'    gf_Query_Old = gf_Query_Old & "                     ELSE "
'    gf_Query_Old = gf_Query_Old & "                       CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN  "
'    gf_Query_Old = gf_Query_Old & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
'    gf_Query_Old = gf_Query_Old & "                          END"
'    gf_Query_Old = gf_Query_Old & "                       ELSE"
'    gf_Query_Old = gf_Query_Old & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
'    gf_Query_Old = gf_Query_Old & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
'    gf_Query_Old = gf_Query_Old & "                          END"
'    gf_Query_Old = gf_Query_Old & "                     END"
'    gf_Query_Old = gf_Query_Old & "                END AS CLASIFICACION2, "
'    gf_Query_Old = gf_Query_Old & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
'    gf_Query_Old = gf_Query_Old & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
'    gf_Query_Old = gf_Query_Old & "                END AS CLAINT2 "
'    gf_Query_Old = gf_Query_Old & "           FROM CRE_HIPCIE "
'    gf_Query_Old = gf_Query_Old & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
'    gf_Query_Old = gf_Query_Old & "          WHERE HIPCIE_PERMES = " & p_MesAnt1
'    gf_Query_Old = gf_Query_Old & "            AND HIPCIE_PERANO = " & p_AnoAnt1
'    gf_Query_Old = gf_Query_Old & "            AND HIPCIE_CLAPRV <> 0 "
'    gf_Query_Old = gf_Query_Old & "          ORDER BY HIPCIE_CLAPRV, HIPCIE_CODPRD) "
'    gf_Query_Old = gf_Query_Old & "  SELECT PRODUCTO, NOMBRECLIENTE, DIASATRASO, TIPOGARAN "
'    gf_Query_Old = gf_Query_Old & "         ,ROUND(VALORGARANTIA,2) AS VALORGARANTIA, TASA "
'    gf_Query_Old = gf_Query_Old & "         ,ROUND(CAPITAL,2) AS CAPITAL, ROUND(CAPITAL_BAL,2) AS CAPITAL_BAL "
'    gf_Query_Old = gf_Query_Old & "         ,ROUND(PROVISION,2) AS PROVISION, CLASIFICACION, CLAINT1 "
'    gf_Query_Old = gf_Query_Old & "         ,ROUND(NVL(PROVISION2,0),2) AS PROVISION2, CLASIFICACION2, CLAINT2 "
'    gf_Query_Old = gf_Query_Old & "         ,ROUND((NVL(PROVISION,0) + NVL(PROV_VOLUNT1,0) - NVL(PROVISION2,0) - NVL(PROV_VOLUNT2,0)),2) AS AJUSTE "
'    gf_Query_Old = gf_Query_Old & "         ,HIPCIE_CLAPRV, HIPCIE_TIPGAR, HIPCIE_TDOCLI, HIPCIE_NDOCLI "
'    gf_Query_Old = gf_Query_Old & "         ,HIPCIE_CLACLI, HIPCIE_CLAALI, PROV_VOLUNT1, PROV_VOLUNT2 "
'    gf_Query_Old = gf_Query_Old & "         ,HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, OPERACION "
'    gf_Query_Old = gf_Query_Old & "   FROM QUERY1 "
'    gf_Query_Old = gf_Query_Old & " LEFT JOIN QUERY2 ON (HIPCIE_NUMOPE = NUMOPE) "
'    gf_Query_Old = gf_Query_Old & "  ORDER BY HIPCIE_CLAPRV, TIPOGARAN, DIASATRASO "
'    gf_Query_Old = gf_Query_Old & "  )"
'End Function

''QUERY 2
Private Function gf_QuerySA() As String
    gf_QuerySA = ""
    gf_QuerySA = gf_QuerySA & " SELECT HIPMAE_NUMOPE, CLAACTUAL, ALIULTMES, HIPMAE_TIPGAR, PRODUCTO, NOMBRECLIENTE, DIASATRASO, "
    gf_QuerySA = gf_QuerySA & "        TIPOGARAN, VALORGARANTIA, CAPITAL, TASA, PROVISION, "
    gf_QuerySA = gf_QuerySA & "        CASE WHEN CLAACTUAL = 0 THEN 'NOR' "
    gf_QuerySA = gf_QuerySA & "             WHEN CLAACTUAL = 1 THEN 'CPP' "
    gf_QuerySA = gf_QuerySA & "             WHEN CLAACTUAL = 2 THEN 'DEF' "
    gf_QuerySA = gf_QuerySA & "             WHEN CLAACTUAL = 3 THEN 'DUD' "
    gf_QuerySA = gf_QuerySA & "             WHEN CLAACTUAL = 4 THEN 'PER' "
    gf_QuerySA = gf_QuerySA & "        END AS CLASIFICACTUAL, "
    gf_QuerySA = gf_QuerySA & "        CASE WHEN ALIULTMES = 0 THEN 'NOR' "
    gf_QuerySA = gf_QuerySA & "             WHEN ALIULTMES = 1 THEN 'CPP' "
    gf_QuerySA = gf_QuerySA & "             WHEN ALIULTMES = 2 THEN 'DEF' "
    gf_QuerySA = gf_QuerySA & "             WHEN ALIULTMES = 3 THEN 'DUD' "
    gf_QuerySA = gf_QuerySA & "             WHEN ALIULTMES = 4 THEN 'PER' "
    gf_QuerySA = gf_QuerySA & "        END AS ALINEAMULTMES"
    gf_QuerySA = gf_QuerySA & "   FROM ( "
    gf_QuerySA = gf_QuerySA & "         WITH CONSULTA AS ("
    gf_QuerySA = gf_QuerySA & "              SELECT A.HIPMAE_NUMOPE AS NUMOPE, "
    gf_QuerySA = gf_QuerySA & "                     (SELECT TIPCLA_CODIGO "
    gf_QuerySA = gf_QuerySA & "                        FROM CTB_TIPCLA "
    gf_QuerySA = gf_QuerySA & "                       WHERE TIPCLA_TIPCRE = '13' "
    gf_QuerySA = gf_QuerySA & "                         AND (SELECT HIPMAE_DIAMOR  FROM CRE_HIPMAE WHERE HIPMAE_NUMOPE =A.HIPMAE_NUMOPE) >= TIPCLA_DIAINI "
    gf_QuerySA = gf_QuerySA & "                         AND (SELECT HIPMAE_DIAMOR  FROM CRE_HIPMAE WHERE HIPMAE_NUMOPE =A.HIPMAE_NUMOPE) <= TIPCLA_DIAFIN ) AS ACTUAL, "
    gf_QuerySA = gf_QuerySA & "                     (SELECT HIPCIE_CLAALI "
    gf_QuerySA = gf_QuerySA & "                        FROM CRE_HIPCIE "
    gf_QuerySA = gf_QuerySA & "                       WHERE HIPCIE_PERMES = " & l_str_PerMes - 1 & " "
    gf_QuerySA = gf_QuerySA & "                         AND HIPCIE_PERANO = " & l_str_PerAno & " "
    gf_QuerySA = gf_QuerySA & "                         AND HIPCIE_NUMOPE = A.HIPMAE_NUMOPE) AS ULTCIE "
    gf_QuerySA = gf_QuerySA & "                FROM CRE_HIPMAE A "
    gf_QuerySA = gf_QuerySA & "               WHERE HIPMAE_DIAMOR > 30 "
    gf_QuerySA = gf_QuerySA & "                 AND HIPMAE_SITUAC = 2 "
    gf_QuerySA = gf_QuerySA & "               ORDER BY HIPMAE_DIAMOR "
    gf_QuerySA = gf_QuerySA & "                          )    "
    gf_QuerySA = gf_QuerySA & "              SELECT HIPMAE_NUMOPE, HIPMAE_TIPGAR, "
    gf_QuerySA = gf_QuerySA & "                     CASE WHEN HIPMAE_CODPRD='001' THEN 'CRC' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='002' THEN 'MIC' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='003' THEN 'CME' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='004' THEN 'MIH' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='006' THEN 'MIC' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='007' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='009' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='010' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='011' THEN 'MIC' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='012' THEN 'UAN' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='013' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='014' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='015' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='016' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='017' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='018' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='019' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='021' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='022' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='023' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='024' THEN 'MIV' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_CODPRD='025' THEN 'MIV' END AS PRODUCTO, "
    gf_QuerySA = gf_QuerySA & "                     TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRECLIENTE, "
    gf_QuerySA = gf_QuerySA & "                     HIPMAE_DIAMOR AS DIASATRASO, "
    gf_QuerySA = gf_QuerySA & "                     CASE WHEN HIPMAE_TIPGAR=1 THEN 'HIP' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_TIPGAR=2 THEN 'BLQ' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_TIPGAR=3 THEN 'FS' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_TIPGAR=4 THEN 'CF' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_TIPGAR=5 THEN 'CP' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_TIPGAR=6 THEN 'RF' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_TIPGAR=8 THEN 'FSHM' "
    gf_QuerySA = gf_QuerySA & "                          WHEN HIPMAE_TIPGAR=9 THEN 'GH' END AS TIPOGARAN, "
    gf_QuerySA = gf_QuerySA & "                     CASE WHEN HIPMAE_TIPGAR IN (1,2) THEN "
    gf_QuerySA = gf_QuerySA & "                          DECODE(HIPMAE_MONGAR,1,HIPMAE_MTOGAR,(HIPMAE_MTOGAR)*HIPMAE_MONEDA) "
    gf_QuerySA = gf_QuerySA & "                          ELSE 0 END AS VALORGARANTIA, "
    gf_QuerySA = gf_QuerySA & "                     DECODE(HIPMAE_MONEDA,1,HIPMAE_SALCAP+HIPMAE_SALCON,(HIPMAE_SALCAP+HIPMAE_SALCON)*HIPMAE_MONEDA) AS CAPITAL, "
    gf_QuerySA = gf_QuerySA & "                     (SELECT TIPPRV_PORCEN  "
    gf_QuerySA = gf_QuerySA & "                        FROM CTB_TIPPRV "
    gf_QuerySA = gf_QuerySA & "                       WHERE TIPPRV_TIPPRV = '2' AND TIPPRV_CLACRE = '13' "
    gf_QuerySA = gf_QuerySA & "                         AND TIPPRV_CLFCRE = ( CASE WHEN (NVL(ULTCIE,0) > ACTUAL) THEN NVL(ULTCIE,0) ELSE TO_NUMBER(ACTUAL) END )"
    gf_QuerySA = gf_QuerySA & "                         AND TIPPRV_CLAGAR = DECODE(HIPMAE_TIPGAR,1,2,1)) AS TASA,  "
    gf_QuerySA = gf_QuerySA & "                "
    gf_QuerySA = gf_QuerySA & "                     ROUND((CASE WHEN (HIPMAE_TIPGAR =1 OR HIPMAE_TIPGAR = 2 ) AND (HIPMAE_CODPRD <> '002' AND HIPMAE_CODPRD <> '011') "
    gf_QuerySA = gf_QuerySA & "                                 THEN ((DECODE(HIPMAE_MONEDA,1,HIPMAE_SALCAP+HIPMAE_SALCON,(HIPMAE_SALCAP+HIPMAE_SALCON)*HIPMAE_MONEDA)) * 2 / 3)"
    gf_QuerySA = gf_QuerySA & "                                      * ((SELECT TipPrv_Porcen  FROM CTB_TIPPRV "
    gf_QuerySA = gf_QuerySA & "                                           WHERE  TIPPRV_TIPPRV = '2' AND TIPPRV_CLACRE = '13' "
    gf_QuerySA = gf_QuerySA & "                                             AND TIPPRV_CLFCRE = ( CASE WHEN (NVL(ULTCIE,0) > ACTUAL) THEN NVL(ULTCIE,0) ELSE TO_NUMBER(ACTUAL) END )"
    gf_QuerySA = gf_QuerySA & "                                             AND TIPPRV_CLAGAR = DECODE(HIPMAE_TIPGAR,1,2,1) )  /100) "
    gf_QuerySA = gf_QuerySA & "                                 ELSE "
    gf_QuerySA = gf_QuerySA & "                                      ((DECODE(HIPMAE_MONEDA,1,HIPMAE_SALCAP+HIPMAE_SALCON,(HIPMAE_SALCAP+HIPMAE_SALCON)*HIPMAE_MONEDA)))"
    gf_QuerySA = gf_QuerySA & "                                      * ((SELECT TipPrv_Porcen  FROM CTB_TIPPRV "
    gf_QuerySA = gf_QuerySA & "                                           WHERE  TIPPRV_TIPPRV = '2' AND TIPPRV_CLACRE = '13' "
    gf_QuerySA = gf_QuerySA & "                                             AND TIPPRV_CLFCRE = ( CASE WHEN (NVL(ULTCIE,0) > ACTUAL) THEN NVL(ULTCIE,0) ELSE TO_NUMBER(ACTUAL) END )"
    gf_QuerySA = gf_QuerySA & "                                             AND TIPPRV_CLAGAR = DECODE(HIPMAE_TIPGAR,1,2,1) ) /100 ) "
    gf_QuerySA = gf_QuerySA & "                             END), 2) AS PROVISION, "
    gf_QuerySA = gf_QuerySA & "                     CASE WHEN (NVL(ULTCIE,0) > ACTUAL) THEN NVL(ULTCIE,0) ELSE TO_NUMBER(ACTUAL) END AS CLAACTUAL, "
    gf_QuerySA = gf_QuerySA & "                     ULTCIE  AS ALIULTMES"
    gf_QuerySA = gf_QuerySA & "                FROM CRE_HIPMAE "
    gf_QuerySA = gf_QuerySA & "               INNER JOIN CLI_DATGEN ON (HIPMAE_TDOCLI = DATGEN_TIPDOC AND HIPMAE_NDOCLI = DATGEN_NUMDOC) "
    gf_QuerySA = gf_QuerySA & "               INNER JOIN CONSULTA   ON (NUMOPE = HIPMAE_NUMOPE)"
    gf_QuerySA = gf_QuerySA & "               WHERE HIPMAE_DIAMOR > 30 "
    gf_QuerySA = gf_QuerySA & "                 AND HIPMAE_SITUAC = 2  "
    gf_QuerySA = gf_QuerySA & "               ORDER BY HIPMAE_DIAMOR "
    gf_QuerySA = gf_QuerySA & "        ) "
End Function

Private Function gf_Query_PrvAct(ByVal p_MesCie As String, ByVal p_AnoCie As String) As String
Dim r_int_MesCieSig     As Integer
Dim r_int_AnoCieSig     As Integer
Dim r_int_DiaAct        As Integer
Dim r_int_UltDia        As Integer
Dim r_int_DifDia        As Integer

   If CInt(p_MesCie) = 12 Then
      r_int_MesCieSig = 1
      r_int_AnoCieSig = CInt(p_AnoCie) + 1
   Else
      r_int_MesCieSig = CInt(p_MesCie) + 1
      r_int_AnoCieSig = CInt(p_AnoCie)
   End If
   
   r_int_DiaAct = Day(date)
   r_int_UltDia = ff_Ultimo_Dia_Mes(r_int_MesCieSig, CInt(r_int_AnoCieSig))
   r_int_DifDia = IIf(r_int_UltDia - r_int_DiaAct > 0, r_int_UltDia - r_int_DiaAct, 0)
   
   gf_Query_PrvAct = ""
   gf_Query_PrvAct = gf_Query_PrvAct & "SELECT HIPMAE_NUMOPE                                  AS OPERACION, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_TDOCLI                                  AS TIPO_DOC, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_NDOCLI                                  AS NUMERO_DOC, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_CODPRD                                  AS COD_PRODUC, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_CODSUB                                  AS COD_SUBPROD, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       ROUND(DECODE(HIPMAE_MONEDA, 1, HIPMAE_SALCAP+HIPMAE_SALCON, (HIPMAE_SALCAP+HIPMAE_SALCON)*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesCie & " AND HIPCIE_PERANO = " & p_AnoCie & ")), 2) AS CAPITAL_SOL, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_SALCAP                                  AS SALDO_TNC, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_SALCON                                  AS SALDO_TC, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_MONEDA                                  AS TIPO_MONEDA, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       (SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE "
   gf_Query_PrvAct = gf_Query_PrvAct & "         WHERE HIPCIE_PERMES = " & p_MesCie & " AND HIPCIE_PERANO = " & p_AnoCie & ")  AS TIPO_CAMBIO, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_TIPGAR                                  AS TIP_GARANT, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_MONGAR                                  AS MONEDA_GARANTIA, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_MTOGAR                                  AS MONTO_GARANTIA, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_FECDES                                  AS FECHA_DESEMBOLSO, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_DIAMOR                                  AS DIAS_ATRA_REAL, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_DIAMOR + " & r_int_DifDia & "           AS DIAS_ATRA_CALC, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_REFINA                                  AS REFINANCIADO, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       HIPMAE_CAPVEN                                  AS CAPITAL_VENCIDO, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       CASE WHEN HIPMAE_CODPRD = '001' THEN 'CRC' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '002' THEN 'MIC' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '003' THEN 'CME' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '004' THEN 'MIH' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '006' THEN 'MIC' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '007' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '009' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '010' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '011' THEN 'MIC' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '012' THEN 'UAN' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '013' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '014' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '015' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '016' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '017' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '018' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '019' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '021' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '022' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '023' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '024' THEN 'MIV' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_CODPRD = '025' THEN 'MIV' END AS PRODUCTO, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       CASE WHEN HIPMAE_TIPGAR = 1 THEN 'HIP' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_TIPGAR = 2 THEN 'HIP' "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_TIPGAR = 3 THEN 'FS'  "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_TIPGAR = 4 THEN 'CF'  "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_TIPGAR = 5 THEN 'CP'  "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_TIPGAR = 6 THEN 'RF'  "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_TIPGAR = 8 THEN 'FSHM' END    AS TIPO_GARANTIA, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       CASE WHEN HIPMAE_TIPGAR IN (1,2) "
   gf_Query_PrvAct = gf_Query_PrvAct & "            THEN ROUND(DECODE(HIPMAE_MONGAR, 1, HIPMAE_MTOGAR, HIPMAE_MTOGAR*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesCie & " AND HIPCIE_PERANO = " & p_AnoCie & ")), 2) "
   gf_Query_PrvAct = gf_Query_PrvAct & "            ELSE 0 "
   gf_Query_PrvAct = gf_Query_PrvAct & "       END                                            AS VALOR_GARANTIA, "
   gf_Query_PrvAct = gf_Query_PrvAct & "       CASE WHEN HIPMAE_DIAMOR + " & r_int_DifDia & " >= 0  AND HIPMAE_DIAMOR + " & r_int_DifDia & " <= 30  THEN 0 "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_DIAMOR + " & r_int_DifDia & " > 30  AND HIPMAE_DIAMOR + " & r_int_DifDia & " <= 60  THEN 1 "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_DIAMOR + " & r_int_DifDia & " > 60  AND HIPMAE_DIAMOR + " & r_int_DifDia & " <= 120 THEN 2 "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_DIAMOR + " & r_int_DifDia & " > 120 AND HIPMAE_DIAMOR + " & r_int_DifDia & " <= 365 THEN 3 "
   gf_Query_PrvAct = gf_Query_PrvAct & "            WHEN HIPMAE_DIAMOR + " & r_int_DifDia & " > 365 THEN 4 "
   gf_Query_PrvAct = gf_Query_PrvAct & "       END                                            AS CLA_INT "
   gf_Query_PrvAct = gf_Query_PrvAct & "  FROM CRE_HIPMAE "
   gf_Query_PrvAct = gf_Query_PrvAct & " INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPMAE_TDOCLI AND DATGEN_NUMDOC = HIPMAE_NDOCLI "
   gf_Query_PrvAct = gf_Query_PrvAct & " WHERE HIPMAE_SITUAC = 2 "
   gf_Query_PrvAct = gf_Query_PrvAct & " ORDER BY PRODUCTO, HIPMAE_DIAMOR "
End Function

Private Function gf_Query_PrvAnt(ByVal p_MesCie As String, ByVal p_AnoCie As String) As String
   gf_Query_PrvAnt = ""
   gf_Query_PrvAnt = gf_Query_PrvAnt & " SELECT HIPCIE_NUMOPE AS OPERACION2, "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "        HIPCIE_CLACLI AS HIPCIE_CLACLI, "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "        HIPCIE_CLAPRV AS HIPCIE_CLAPRV, "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "        ROUND(DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*(SELECT MAX(HIPCIE_TIPCAM) FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_MesCie & " AND HIPCIE_PERANO = " & p_AnoCie & ")),2) AS PROVISION2, "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "        CASE WHEN HIPCIE_FLGREF = 1 THEN "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                     WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                     WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                     WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                END "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "             ELSE "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "               CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN  "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                  CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                       WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                       WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                       WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                  END "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "               ELSE "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                  CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                       WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                       WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                       WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "                  END "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "             END "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "        END AS CLA_ALI2, "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "        CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "             WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "             WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "             WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "             WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "        END AS CLA_INT2 "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "   FROM CRE_HIPCIE "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "  WHERE HIPCIE_PERMES = " & p_MesCie & " "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "    AND HIPCIE_PERANO = " & p_AnoCie & " "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "    AND HIPCIE_CLAPRV <> 0 "
   gf_Query_PrvAnt = gf_Query_PrvAnt & "  ORDER BY HIPCIE_CLAPRV, HIPCIE_CODPRD   "
End Function

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub

Private Function fs_Calcula_CuotasPagadasConAtraso(ByVal p_NumOpe As String, ByVal p_Ano As Integer, ByVal p_Mes As Integer) As String
Dim r_str_Cadena     As String
Dim r_rst_Cuotas     As ADODB.Recordset
Dim r_str_FecPro     As String
Dim r_int_TotCuo     As Integer
Dim r_int_CuoAtr     As Integer
Dim r_int_DifDia     As Integer

   fs_Calcula_CuotasPagadasConAtraso = ""
   r_str_FecPro = Format(p_Ano, "0000") & Format(p_Mes, "00") & Format(ff_Ultimo_Dia_Mes(p_Mes, p_Ano), "00")
   
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT CUOCIE_NUMCUO AS NUM_CUOTA, "
   r_str_Cadena = r_str_Cadena & "       CUOCIE_SITUAC AS SITUACION, "
   r_str_Cadena = r_str_Cadena & "       CUOCIE_FECVCT AS FEC_VCTO, "
   r_str_Cadena = r_str_Cadena & "       CUOCIE_FECPAG AS FEC_PAGO, "
   r_str_Cadena = r_str_Cadena & "       DECODE(CUOCIE_FECPAG, 0, 0, TO_DATE(CUOCIE_FECPAG,'YYYY/MM/DD') - TO_DATE(CUOCIE_FECVCT,'YYYY/MM/DD')) AS DIF_DIAS "
   r_str_Cadena = r_str_Cadena & "  FROM CRE_CUOCIE "
   r_str_Cadena = r_str_Cadena & " WHERE CUOCIE_PERMES = '" & p_Mes & "' "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_PERANO = '" & p_Ano & "' "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_NUMOPE = '" & p_NumOpe & "' "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_TIPCRO = 1 "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_FECVCT < " & r_str_FecPro & " "
   r_str_Cadena = r_str_Cadena & "ORDER BY CUOCIE_NUMCUO "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 3) Then
      Exit Function
   End If
    
   If (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
      Exit Function
   End If
   
   r_int_TotCuo = 0
   r_int_CuoAtr = 0
   r_rst_Cuotas.MoveFirst
   Do While Not r_rst_Cuotas.EOF
      r_int_TotCuo = r_int_TotCuo + 1
      If r_rst_Cuotas!SITUACION = 1 Then
         If r_rst_Cuotas!DIF_DIAS > 30 Then
            r_int_CuoAtr = r_int_CuoAtr + 1
         End If
      Else
         r_int_DifDia = DateDiff("d", Format(gf_FormatoFecha(Trim(r_rst_Cuotas!FEC_VCTO)), "dd/mm/yyyy"), Format(gf_FormatoFecha(Trim(r_str_FecPro)), "dd/mm/yyyy"), vbMonday)
         If r_int_DifDia > 30 Then
            r_int_CuoAtr = r_int_CuoAtr + 1
         End If
      End If
      r_rst_Cuotas.MoveNext
   Loop
    
   fs_Calcula_CuotasPagadasConAtraso = CStr(r_int_CuoAtr) & "/" & CStr(r_int_TotCuo)
   r_rst_Cuotas.Close
   Set r_rst_Cuotas = Nothing
End Function

Private Sub fs_ObtieneDatosRCC(ByVal p_TipDoc As String, ByVal p_NumDoc As String, p_NumMes As String, ByVal p_NumAno As String, ByRef p_NumEnt As Integer, ByRef p_TotDeu As Double)
Dim r_str_Cadena     As String
Dim r_rst_DatRcc     As ADODB.Recordset

   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT RCCCAB_NUMEMP AS NUM_ENT, "
   r_str_Cadena = r_str_Cadena & "       RCCCAB_DEUCA0+RCCCAB_DEUCA1+RCCCAB_DEUCA2+RCCCAB_DEUCA3+RCCCAB_DEUCA4 AS TOT_DEU "
   r_str_Cadena = r_str_Cadena & "  FROM CLI_RCCCAB "
   r_str_Cadena = r_str_Cadena & " WHERE RCCCAB_TIPDOC = " & Trim(p_TipDoc) & " "
   r_str_Cadena = r_str_Cadena & "   AND TRIM(RCCCAB_NUMDOC) = '" & Trim(p_NumDoc) & "' "
   r_str_Cadena = r_str_Cadena & "   AND RCCCAB_PERMES = " & Trim(p_NumMes) & " "
   r_str_Cadena = r_str_Cadena & "   AND RCCCAB_PERANO = " & Trim(p_NumAno) & " "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_DatRcc, 3) Then
      Exit Sub
   End If
      
   If Not (r_rst_DatRcc.BOF And r_rst_DatRcc.EOF) Then
      r_rst_DatRcc.MoveFirst
      p_NumEnt = r_rst_DatRcc!NUM_ENT
      p_TotDeu = r_rst_DatRcc!TOT_DEU
   End If
   
   r_rst_DatRcc.Close
   Set r_rst_DatRcc = Nothing
End Sub

Private Function fs_MuestraTipo(ByVal p_PrvAct As String, ByVal p_PrvAnt As Variant) As String
   fs_MuestraTipo = ""
   
   If p_PrvAct = "NOR" Then
      fs_MuestraTipo = "NORMAL"
      Exit Function
   End If
   If IsNull(p_PrvAnt) Or p_PrvAnt = "" Then
      fs_MuestraTipo = "NUEVO"
      If InStr(1, p_PrvAct, "ALI", vbTextCompare) > 0 Then
         fs_MuestraTipo = "ALINEADO"
      End If
      Exit Function
   End If
   If p_PrvAct = p_PrvAnt Then
      fs_MuestraTipo = "IGUAL"
      If InStr(1, p_PrvAct, "ALI", vbTextCompare) > 0 Then
         fs_MuestraTipo = "ALINEADO"
      End If
   Else
      If InStr(1, p_PrvAct, "ALI", vbTextCompare) > 0 Then
         fs_MuestraTipo = "ALINEADO"
      Else
         If p_PrvAct = "CPP" Then
            fs_MuestraTipo = "MEJOR"
         End If
         If p_PrvAct = "DEF" Then
            If p_PrvAnt = "CPP" Then
               fs_MuestraTipo = "PEOR"
            Else
               fs_MuestraTipo = "MEJOR"
            End If
         End If
         If p_PrvAct = "DUD" Then
            If p_PrvAnt = "DEF" Or p_PrvAnt = "CPP" Then
               fs_MuestraTipo = "PEOR"
            Else
               fs_MuestraTipo = "MEJOR"
            End If
         End If
         If p_PrvAct = "PER" Then
            If p_PrvAnt = "DEF" Or p_PrvAnt = "DUD" Or p_PrvAnt = "CPP" Then
               fs_MuestraTipo = "PEOR"
            Else
               fs_MuestraTipo = "IGUAL"
            End If
         End If
         If p_PrvAnt = "PER" Then
            fs_MuestraTipo = "MEJOR"
         End If
      End If
   End If
End Function

Private Function fs_CuotasAtrasadas_Cierre(ByVal p_NumOpe As String, ByVal p_Ano As Integer, ByVal p_Mes As Integer) As Integer
Dim r_str_Cadena     As String
Dim r_rst_Cuotas     As ADODB.Recordset
Dim r_str_FecPro     As String
   
   fs_CuotasAtrasadas_Cierre = 0
   r_str_FecPro = Format(p_Ano, "0000") & Format(p_Mes, "00") & Format(ff_Ultimo_Dia_Mes(p_Mes, p_Ano), "00")
   
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS CUOTAS_ATRASADAS"
   r_str_Cadena = r_str_Cadena & "  FROM CRE_CUOCIE "
   r_str_Cadena = r_str_Cadena & " WHERE CUOCIE_PERMES = '" & p_Mes & "' "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_PERANO = '" & p_Ano & "' "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_NUMOPE = '" & p_NumOpe & "' "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_TIPCRO = 1 "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_SITUAC = 2 "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_FECVCT < " & r_str_FecPro & " "
   r_str_Cadena = r_str_Cadena & "ORDER BY CUOCIE_NUMCUO "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 3) Then
      Exit Function
   End If
    
   If (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
      Exit Function
   End If
   
   fs_CuotasAtrasadas_Cierre = r_rst_Cuotas!CUOTAS_ATRASADAS
   
   r_rst_Cuotas.Close
   Set r_rst_Cuotas = Nothing
End Function

Private Function fs_Clasifica_Dudoso_Perdida(ByVal p_NumOpe As String, ByVal p_Ano As Integer, ByVal p_Mes As Integer) As String
Dim r_str_Cadena     As String
Dim r_rst_DudPer     As ADODB.Recordset
Dim r_arr_Genera()   As moddat_tpo_Genera
Dim r_int_Contad     As Integer
Dim r_int_ConDud     As Integer
Dim r_int_ConPer     As Integer
Dim r_bol_FlgAux     As Boolean
   
   r_int_ConDud = 0
   r_int_ConPer = 0
   r_bol_FlgAux = False
   fs_Clasifica_Dudoso_Perdida = ""
   ReDim r_arr_Genera(0)
   
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & " SELECT CASE WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
   r_str_Cadena = r_str_Cadena & "             WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
   r_str_Cadena = r_str_Cadena & "         END                    AS CLASIFICACION, "
   r_str_Cadena = r_str_Cadena & "         HIPCIE_PERMES AS MES, HIPCIE_PERANO AS ANIO "
   r_str_Cadena = r_str_Cadena & "   FROM CRE_HIPCIE "
   r_str_Cadena = r_str_Cadena & "  WHERE HIPCIE_PERMES > 0 "
   r_str_Cadena = r_str_Cadena & "    AND HIPCIE_PERANO <= " & p_Ano & " "
   r_str_Cadena = r_str_Cadena & "    AND HIPCIE_NUMOPE = '" & p_NumOpe & "' "
   r_str_Cadena = r_str_Cadena & "    AND HIPCIE_CLACLI IN (3,4) "
   r_str_Cadena = r_str_Cadena & "  ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_DudPer, 3) Then
      Exit Function
   End If
    
   If (r_rst_DudPer.BOF And r_rst_DudPer.EOF) Then
      Exit Function
   End If
   
   Do While Not r_rst_DudPer.EOF
      If p_Mes = r_rst_DudPer!Mes And p_Ano = r_rst_DudPer!ANIO Then
         r_bol_FlgAux = True
      End If
      
      If r_bol_FlgAux = True Then
         ReDim Preserve r_arr_Genera(UBound(r_arr_Genera) + 1)
         r_arr_Genera(UBound(r_arr_Genera)).Genera_Codigo = r_rst_DudPer!CLASIFICACION
         r_arr_Genera(UBound(r_arr_Genera)).Genera_CodIns = r_rst_DudPer!Mes
         r_arr_Genera(UBound(r_arr_Genera)).Genera_ConHip = r_rst_DudPer!ANIO
      End If
      r_rst_DudPer.MoveNext
   Loop
   
   r_bol_FlgAux = False
   If UBound(r_arr_Genera) > 0 Then
      If r_arr_Genera(1).Genera_Codigo = "DUD" Then
         r_int_ConDud = 1
         For r_int_Contad = 2 To UBound(r_arr_Genera)
            If r_arr_Genera(r_int_Contad).Genera_Codigo = "DUD" And (IIf(r_arr_Genera(r_int_Contad - 1).Genera_CodIns = 1, 13, r_arr_Genera(r_int_Contad - 1).Genera_CodIns) - r_arr_Genera(r_int_Contad).Genera_CodIns) = 1 Then
               r_int_ConDud = r_int_ConDud + 1
            Else
               r_bol_FlgAux = True
            End If
            
            If r_bol_FlgAux = True Then
               fs_Clasifica_Dudoso_Perdida = r_int_ConDud & "D"
               r_rst_DudPer.Close
               Set r_rst_DudPer = Nothing
               Exit Function
            End If
         Next r_int_Contad
         
      ElseIf r_arr_Genera(1).Genera_Codigo = "PER" Then
         r_int_ConPer = 1
         
         For r_int_Contad = 2 To UBound(r_arr_Genera)
            If r_arr_Genera(r_int_Contad).Genera_Codigo = "PER" And (IIf(r_arr_Genera(r_int_Contad - 1).Genera_CodIns = 1, 13, r_arr_Genera(r_int_Contad - 1).Genera_CodIns) - r_arr_Genera(r_int_Contad).Genera_CodIns) = 1 Then
               r_int_ConPer = r_int_ConPer + 1
            Else
               r_bol_FlgAux = True
            End If
            
            If r_bol_FlgAux = True Then
               fs_Clasifica_Dudoso_Perdida = r_int_ConPer & "P"
               r_rst_DudPer.Close
               Set r_rst_DudPer = Nothing
               Exit Function
            End If
         Next r_int_Contad
      End If
      
      If r_arr_Genera(1).Genera_Codigo = "DUD" Then
         fs_Clasifica_Dudoso_Perdida = r_int_ConDud & "D"
      ElseIf r_arr_Genera(1).Genera_Codigo = "PER" Then
         fs_Clasifica_Dudoso_Perdida = r_int_ConPer & "P"
      End If
   End If
   
   r_rst_DudPer.Close
   Set r_rst_DudPer = Nothing
End Function

Private Function fs_Busca_Excepciones(ByVal p_NumOpe As String) As String
Dim r_str_Cadena     As String
Dim r_rst_CodExc     As ADODB.Recordset
   
   fs_Busca_Excepciones = ""

   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT SEGEXC_MOTEXC "
   r_str_Cadena = r_str_Cadena & "  FROM CRE_HIPMAE "
   r_str_Cadena = r_str_Cadena & " INNER JOIN TRA_SEGEXC ON SEGEXC_NUMSOL = HIPMAE_NUMSOL AND SEGEXC_CODINS = 21 "
   r_str_Cadena = r_str_Cadena & " WHERE HIPMAE_NUMOPE = '" & p_NumOpe & "' "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_CodExc, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_CodExc.BOF And r_rst_CodExc.EOF) Then
      r_rst_CodExc.MoveFirst
      Do While Not r_rst_CodExc.EOF
         If fs_Busca_Excepciones = "" Then
            fs_Busca_Excepciones = CStr(r_rst_CodExc!SEGEXC_MOTEXC)
         Else
            fs_Busca_Excepciones = fs_Busca_Excepciones & " , " & CStr(r_rst_CodExc!SEGEXC_MOTEXC)
         End If
         r_rst_CodExc.MoveNext
      Loop
   End If
   
   r_rst_CodExc.Close
   Set r_rst_CodExc = Nothing
End Function

