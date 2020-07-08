VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptCtb_26 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   6720
   ClientTop       =   6675
   ClientWidth     =   6045
   Icon            =   "GesCtb_frm_852.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2805
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6045
      _Version        =   65536
      _ExtentX        =   10663
      _ExtentY        =   4948
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
         Top             =   30
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
            Caption         =   "Reporte de Cobranzas a Clientes Morosos y Alienados"
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
            Picture         =   "GesCtb_frm_852.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   750
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
            Picture         =   "GesCtb_frm_852.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5340
            Picture         =   "GesCtb_frm_852.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1305
         Left            =   30
         TabIndex        =   9
         Top             =   1440
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
         _ExtentY        =   2302
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
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   165
            Width           =   2415
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1650
            TabIndex        =   1
            Top             =   480
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDateTime ipp_FecBus 
            Height          =   315
            Left            =   1650
            TabIndex        =   2
            Top             =   810
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
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
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha de consulta:"
            Height          =   285
            Left            =   90
            TabIndex        =   12
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "Mes de Proceso:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   195
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Año de Proceso:"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   525
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_PerMes        As String
Dim l_str_PerAno        As String
Dim l_str_MesAnt        As String
Dim l_str_AnoAnt        As String

Private Sub cmd_ExpExc_Click()
Dim r_dat_FecMin        As Date
Dim r_dat_FecMax        As Date
Dim r_str_MesSig        As String
Dim r_str_AnoSig        As String

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
   If Not IsDate(ipp_FecBus) Then
      MsgBox "La fecha ingresada no es valida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecBus)
      Exit Sub
   End If
   
   r_str_MesSig = IIf(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 12, 1, cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1)
   r_str_AnoSig = IIf(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 12, CLng(ipp_PerAno.Text) + 1, ipp_PerAno.Text)
   r_dat_FecMin = CDate("01" & "/" & r_str_MesSig & "/" & r_str_AnoSig)
   r_dat_FecMax = moddat_g_str_FecSis
   
   If CDate(ipp_FecBus.Text) < r_dat_FecMin Then
      MsgBox "La fecha ingresada no puede ser menor a " & CStr(r_dat_FecMin), vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecBus)
      Exit Sub
   End If
   If CDate(ipp_FecBus.Text) > r_dat_FecMax Then
      MsgBox "La fecha ingresada no puede ser mayor a " & CStr(r_dat_FecMax), vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecBus)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   l_str_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   l_str_PerAno = ipp_PerAno.Text
   l_str_MesAnt = IIf(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 1, 12, cmb_PerMes.ItemData(cmb_PerMes.ListIndex) - 1)
   l_str_AnoAnt = IIf(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 1, CLng(ipp_PerAno.Text) - 1, ipp_PerAno.Text)
   
   Screen.MousePointer = 11
   Me.Enabled = False
   Call fs_GenExc(l_str_PerMes, l_str_PerAno, l_str_MesAnt, l_str_AnoAnt, Format(ipp_FecBus.Text, "YYYYMMDD"))
   Me.Enabled = True
   Screen.MousePointer = 0
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
    ipp_FecBus.Text = date
End Sub

Private Sub fs_GenExc(ByVal p_MesAct As Integer, ByVal p_AnoAct As Integer, ByVal p_MesAnt As Integer, ByVal p_AnoAnt As Integer, ByVal p_FecBus As String)
Dim r_obj_Excel     As Excel.Application
Dim r_int_nrofil    As Integer
Dim r_int_Nindex    As Integer
Dim r_int_nroaux    As Integer
Dim r_bol_FlagOp    As Boolean
Dim r_dbl_SumCap    As Double
Dim r_dbl_SumPrv    As Double
    
   r_dbl_SumCap = 0
   r_dbl_SumPrv = 0
   r_int_Nindex = 1
   
   r_bol_FlagOp = False
   g_str_Parame = gf_Query(p_MesAct, p_AnoAct, p_MesAnt, p_AnoAnt)
     
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
   r_int_nrofil = 2
    
   With r_obj_Excel.ActiveSheet
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").Merge
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").Font.Bold = True
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").Font.Size = 14
      .Range("A" & r_int_nrofil & ":L" & r_int_nrofil & "").Font.Underline = xlUnderlineStyleSingle
      
      .Cells(r_int_nrofil, 1) = "CLIENTES MOROSOS Y ALINEADOS AL MES DE " & Trim(cmb_PerMes.Text) & " DEL " & Trim(ipp_PerAno.Text) & " - REGISTRO DE COBRANZAS"
      r_int_nrofil = r_int_nrofil + 2
      
      .Range("H" & r_int_nrofil & ":J" & r_int_nrofil & "").Merge
      .Range("H" & r_int_nrofil & ":J" & r_int_nrofil & "").HorizontalAlignment = xlHAlignCenter
      .Range("K" & r_int_nrofil & ":M" & r_int_nrofil & "").HorizontalAlignment = xlHAlignCenter
      .Range("H" & r_int_nrofil & ":J" & r_int_nrofil & "").Font.Bold = True
      .Cells(r_int_nrofil, 8) = "ACTUAL"
      
      r_int_nrofil = r_int_nrofil + 1
      .Cells(r_int_nrofil, 1) = "PRODUCTO"
      .Cells(r_int_nrofil, 2) = "NOMBRE DEL CLIENTE"
      .Cells(r_int_nrofil, 3) = "DÍAS ATRASO"
      .Cells(r_int_nrofil, 4) = "GARANTÍA"
      .Cells(r_int_nrofil, 5) = "VAL. GARANTÍA"
      .Cells(r_int_nrofil, 6) = "CAPITAL"
      .Cells(r_int_nrofil, 7) = "TASA %"
      .Cells(r_int_nrofil, 8) = "PROVISIÓN"
      .Cells(r_int_nrofil, 9) = "CLA. ALINEADA"
      .Cells(r_int_nrofil, 10) = "CLA. INTERNA"
      .Cells(r_int_nrofil, 11) = "GESTION DE COBRANZAS"
      .Cells(r_int_nrofil, 12) = "CODIGO DE EXCEPCION"
      
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 12)).Font.Bold = True
      .Range(.Cells(r_int_nrofil, 1), .Cells(r_int_nrofil, 12)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 11
      .Columns("B").ColumnWidth = 46
      .Columns("C").ColumnWidth = 12
      .Columns("D").ColumnWidth = 12
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 12
      .Columns("G").ColumnWidth = 8
      .Columns("H").ColumnWidth = 15
      .Columns("I").ColumnWidth = 14
      .Columns("J").ColumnWidth = 13
      .Columns("K").ColumnWidth = 50
      .Columns("L").ColumnWidth = 22
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("L").HorizontalAlignment = xlHAlignCenter
   End With
   
   Do While r_int_Nindex <= 4
      r_int_nroaux = 1
     
      Do While r_int_nroaux < 3
         r_dbl_SumCap = 0
         r_dbl_SumPrv = 0
         g_rst_Princi.Filter = adFilterNone
         g_rst_Princi.MoveFirst
         r_int_nrofil = r_int_nrofil + 1
         
         If r_int_nroaux = 1 Then
            g_rst_Princi.Filter = "HIPCIE_CLAPRV = " & r_int_Nindex & " AND HIPCIE_TIPGAR > 2"
         Else
            g_rst_Princi.Filter = "HIPCIE_CLAPRV = " & r_int_Nindex & " AND HIPCIE_TIPGAR < 3"
         End If
         
         If g_rst_Princi.EOF Then
            r_bol_FlagOp = True
         End If
          
         Do While Not g_rst_Princi.EOF
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 1) = Trim(g_rst_Princi!PRODUCTO)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 2) = Trim(g_rst_Princi!NOMBRECLIENTE)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 3) = Format(g_rst_Princi!DIASATRASO, "#00")
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 3), r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 3)).NumberFormat = "#,##0"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 4) = Trim(g_rst_Princi!TIPOGARAN)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 5) = CDbl(Format(g_rst_Princi!VALORGARANTIA, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 5), r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 5)).NumberFormat = "###,###,##0.00"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 6) = CDbl(Format(g_rst_Princi!CAPITAL, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 6), r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 6)).NumberFormat = "###,###,##0.00"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = CDbl(Format(g_rst_Princi!TASA, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7), r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7)).NumberFormat = "###,###,##0.00"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(Format(g_rst_Princi!PROVISION, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8), r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8)).NumberFormat = "###,###,##0.00"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = Trim(g_rst_Princi!CLASIFICACION)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 10) = Trim(g_rst_Princi!CLAINT1)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 11) = fs_Busca_Cobranzas(g_rst_Princi!OPERACION, p_AnoAct, p_MesAct, p_FecBus)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 12) = fs_Busca_Excepciones(g_rst_Princi!OPERACION)
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 12), r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 12)).NumberFormat = "#,##0"
            
            r_dbl_SumCap = r_dbl_SumCap + CDbl(Format(g_rst_Princi!CAPITAL, "###,###,##0.00"))
            r_dbl_SumPrv = r_dbl_SumPrv + CDbl(Format(g_rst_Princi!PROVISION, "###,###,##0.00"))
            r_int_nrofil = r_int_nrofil + 1
            g_rst_Princi.MoveNext
            DoEvents
         Loop
          
         If r_dbl_SumCap = 0 Then
            r_int_nrofil = r_int_nrofil - 1
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 6) = CDbl(Format(r_dbl_SumCap, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(Format(r_dbl_SumPrv, "###,###,##0.00"))
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 6).Font.Bold = True
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8).Font.Bold = True
         End If
         r_int_nroaux = r_int_nroaux + 1
      Loop
      
      If Not r_bol_FlagOp Then r_int_nrofil = r_int_nrofil + 2
      r_bol_FlagOp = False
      r_int_Nindex = r_int_Nindex + 1
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Function gf_Query(ByVal p_MesAct1 As Integer, ByVal p_AnoAct1 As Integer, ByVal p_MesAnt1 As Integer, ByVal p_AnoAnt1 As Integer) As String
    gf_Query = ""
    gf_Query = gf_Query & " SELECT * FROM ( "
    gf_Query = gf_Query & " WITH QUERY1 AS ( "
    gf_Query = gf_Query & "         SELECT HIPCIE_NUMOPE, HIPCIE_TDOCLI, HIPCIE_NDOCLI, HIPCIE_CLACLI, HIPCIE_NUMOPE AS OPERACION, "
    gf_Query = gf_Query & "                HIPCIE_CLAALI, HIPCIE_CLAPRV, HIPCIE_TIPGAR, HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, "
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
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=2 THEN 'HIP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=3 THEN 'FS'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=4 THEN 'CF'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=5 THEN 'CP'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=6 THEN 'RF'  "
    gf_Query = gf_Query & "                     WHEN HIPCIE_TIPGAR=8 THEN 'FSHM' END AS TIPOGARAN, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_TIPGAR IN (1,2) THEN DECODE(HIPCIE_MONGAR,1,HIPCIE_MTOGAR,(HIPCIE_MTOGAR)*HIPCIE_TIPCAM) "
    gf_Query = gf_Query & "                     ELSE 0 END AS VALORGARANTIA, "
    gf_Query = gf_Query & "                DECODE(HIPCIE_TIPMON,1,HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF,(HIPCIE_SALCAP+HIPCIE_SALCON-HIPCIE_INTDIF)*HIPCIE_TIPCAM) AS CAPITAL, "
    gf_Query = gf_Query & "                (SELECT TIPPRV_PORCEN "
    gf_Query = gf_Query & "                   FROM CTB_TIPPRV "
    gf_Query = gf_Query & "                  WHERE TIPPRV_TIPPRV = '2' "
    gf_Query = gf_Query & "                    AND TIPPRV_CLACRE = '13' "
    gf_Query = gf_Query & "                    AND TIPPRV_CLFCRE = HIPCIE_CLAPRV "
    gf_Query = gf_Query & "                    AND TIPPRV_CLAGAR = DECODE(HIPCIE_TIPGAR,1,2,1)) AS TASA, "
    gf_Query = gf_Query & "                DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*HIPCIE_TIPCAM) AS PROVISION, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_FLGREF = 1 THEN "
    gf_Query = gf_Query & "                        CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-REF' "
    gf_Query = gf_Query & "                             WHEN HIPCIE_CLAPRV = 4 THEN 'PER-REF' "
    gf_Query = gf_Query & "                        END "
    gf_Query = gf_Query & "                     ELSE "
    gf_Query = gf_Query & "                       CASE WHEN HIPCIE_CLACLI = HIPCIE_CLAALI THEN "
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER' "
    gf_Query = gf_Query & "                          END "
    gf_Query = gf_Query & "                       ELSE "
    gf_Query = gf_Query & "                          CASE WHEN HIPCIE_CLAPRV = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 2 THEN 'DEF-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 3 THEN 'DUD-ALI' "
    gf_Query = gf_Query & "                               WHEN HIPCIE_CLAPRV = 4 THEN 'PER-ALI' "
    gf_Query = gf_Query & "                          END "
    gf_Query = gf_Query & "                     END "
    gf_Query = gf_Query & "                END AS CLASIFICACION, "
    gf_Query = gf_Query & "                CASE WHEN HIPCIE_CLACLI = 0 THEN 'NOR' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 1 THEN 'CPP' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 2 THEN 'DEF' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 3 THEN 'DUD' "
    gf_Query = gf_Query & "                     WHEN HIPCIE_CLACLI = 4 THEN 'PER' "
    gf_Query = gf_Query & "                END AS CLAINT1 "
    gf_Query = gf_Query & "           FROM CRE_HIPCIE "
    gf_Query = gf_Query & "          INNER JOIN CLI_DATGEN ON (HIPCIE_TDOCLI = DATGEN_TIPDOC AND HIPCIE_NDOCLI = DATGEN_NUMDOC) "
    gf_Query = gf_Query & "          WHERE HIPCIE_PERMES = " & p_MesAct1
    gf_Query = gf_Query & "            AND HIPCIE_PERANO = " & p_AnoAct1
    gf_Query = gf_Query & "            AND HIPCIE_CLAPRV <> 0"
    gf_Query = gf_Query & "          ORDER BY HIPCIE_CLAPRV, TIPOGARAN, DIASATRASO), "
    gf_Query = gf_Query & " QUERY2 AS ("
    gf_Query = gf_Query & "         SELECT HIPCIE_NUMOPE AS NUMOPE, "
    gf_Query = gf_Query & "                DECODE(HIPCIE_TIPMON, 1, HIPCIE_PRVESP, HIPCIE_PRVESP*HIPCIE_TIPCAM) AS PROVISION2, "
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
    gf_Query = gf_Query & "          ORDER BY HIPCIE_CLAPRV, HIPCIE_CODPRD) "
    gf_Query = gf_Query & "  SELECT PRODUCTO, NOMBRECLIENTE, DIASATRASO, TIPOGARAN "
    gf_Query = gf_Query & "         ,ROUND(VALORGARANTIA,2) AS VALORGARANTIA "
    gf_Query = gf_Query & "         ,ROUND(CAPITAL,2) AS CAPITAL, TASA "
    gf_Query = gf_Query & "         ,ROUND(PROVISION,2) AS PROVISION, CLASIFICACION,CLAINT1 "
    gf_Query = gf_Query & "         ,ROUND(NVL(PROVISION2,0),2) AS PROVISION2, CLASIFICACION2,CLAINT2 "
    gf_Query = gf_Query & "         ,ROUND((PROVISION - NVL(PROVISION2,0)),2) AS AJUSTE "
    gf_Query = gf_Query & "         ,HIPCIE_TIPGAR, HIPCIE_TDOCLI, HIPCIE_NDOCLI "
    gf_Query = gf_Query & "         ,HIPCIE_CLAPRV, HIPCIE_CLACLI, HIPCIE_CLAALI "
    gf_Query = gf_Query & "         ,HIPCIE_CBRFMV, HIPCIE_CBRFMV_RC, OPERACION  "
    gf_Query = gf_Query & "   FROM QUERY1 "
    gf_Query = gf_Query & "   LEFT JOIN QUERY2 ON (HIPCIE_NUMOPE = NUMOPE) "
    gf_Query = gf_Query & "  ORDER BY HIPCIE_CLAPRV, TIPOGARAN, DIASATRASO "
    gf_Query = gf_Query & "  )"
     
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

Private Function fs_Busca_Cobranzas(ByVal p_NumOpe As String, ByVal p_Ano As Integer, ByVal p_Mes As Integer, ByVal p_FecBus As String) As String
Dim r_str_Cadena     As String
Dim r_rst_Cuotas     As ADODB.Recordset
Dim r_str_FecPro     As String
Dim r_int_CuoVen     As Integer
Dim r_int_CuoPag     As Integer

   fs_Busca_Cobranzas = ""
   r_str_FecPro = Format(p_Ano, "0000") & Format(p_Mes, "00") & Format(ff_Ultimo_Dia_Mes(p_Mes, p_Ano), "00")
   
   'Cuotas Vencidas
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS CUO_VENCIDAS "
   r_str_Cadena = r_str_Cadena & "  FROM CRE_CUOCIE "
   r_str_Cadena = r_str_Cadena & " WHERE CUOCIE_PERMES = '" & p_Mes & "' "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_PERANO = '" & p_Ano & "' "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_NUMOPE = '" & p_NumOpe & "' "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_TIPCRO = 1 "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_SITUAC = 2 "
   r_str_Cadena = r_str_Cadena & "   AND CUOCIE_FECVCT < " & r_str_FecPro & " "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 3) Then
      Exit Function
   End If
    
   If (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
      Exit Function
   End If
   
   r_rst_Cuotas.MoveFirst
   r_int_CuoVen = r_rst_Cuotas!CUO_VENCIDAS
   
   r_rst_Cuotas.Close
   Set r_rst_Cuotas = Nothing
   
   'Cuotas Pagadas
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS CUO_PAGADAS "
   r_str_Cadena = r_str_Cadena & "  FROM CRE_HIPCUO "
   r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 1 "
   r_str_Cadena = r_str_Cadena & "   AND HIPCUO_SITUAC = 1 "
   r_str_Cadena = r_str_Cadena & "   AND HIPCUO_FECPAG > " & r_str_FecPro & " "
   r_str_Cadena = r_str_Cadena & "   AND HIPCUO_FECPAG <= " & p_FecBus & " "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 3) Then
      Exit Function
   End If
    
   If (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
      Exit Function
   End If
   
   r_rst_Cuotas.MoveFirst
   r_int_CuoPag = r_rst_Cuotas!CUO_PAGADAS
   
   r_rst_Cuotas.Close
   Set r_rst_Cuotas = Nothing
   
   fs_Busca_Cobranzas = CStr(r_int_CuoPag) & "  CUOTA(S) PAGADA(S)   DE   " & CStr(r_int_CuoVen) & "  CUOTA(S) VENCIDA(S)"
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
