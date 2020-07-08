VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2700
   ClientLeft      =   10680
   ClientTop       =   2835
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   7260
   Begin Threed.SSPanel SSPanel1 
      Height          =   2715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      _Version        =   65536
      _ExtentX        =   12832
      _ExtentY        =   4789
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
            Top             =   315
            Width           =   3495
            _Version        =   65536
            _ExtentX        =   6165
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Método de Indicardor Básico"
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
            Picture         =   "GesCtb_frm_178.frx":0000
            Top             =   60
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
            Picture         =   "GesCtb_frm_178.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_178.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_178.frx":0B8E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   2370
            Top             =   60
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
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   4110
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   2805
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1170
            TabIndex        =   10
            Top             =   240
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1170
            TabIndex        =   11
            Top             =   600
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
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   150
            TabIndex        =   14
            Top             =   660
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   150
            TabIndex        =   13
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   3060
            TabIndex        =   12
            Top             =   300
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ExpExc_Click()
   Screen.MousePointer = 11
   Call fs_prueba
   Screen.MousePointer = 0
   
End Sub


Private Sub fs_prueba()
   Dim r_obj_Excel    As Excel.Application
   Dim r_int_NroFil   As Integer
   Dim r_int_NroIni   As Integer
   Dim r_int_ColumA   As Integer
   Dim r_int_ColumB   As Integer
   Dim r_int_ColumC   As Integer
   Dim r_int_ColumD   As Integer
   Dim r_int_ColumE   As Integer
   Dim r_lng_Import   As Long
   
   r_int_ColumA = 1
   r_int_ColumB = 2
   r_int_ColumC = 3
   r_int_ColumD = 4
   r_int_ColumE = 5
   r_lng_Import = 0

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM TT_BALANCE_COMPROB_PRUEBA "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
      Exit Sub
   End If

   g_str_Parame = ""
   g_str_Parame = "usp_balance_comprob_prueba ("
   g_str_Parame = g_str_Parame & "'001', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'" & ipp_FecIni.Text & "', "
   g_str_Parame = g_str_Parame & "'" & ipp_FecIni.Text & "', "
   g_str_Parame = g_str_Parame & "'', ''"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      moddat_g_int_CntErr = moddat_g_int_CntErr + 1
   Else
      moddat_g_int_FlgGOK = True
   End If

   If moddat_g_int_CntErr = 6 Then
      If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
         Exit Sub
      Else
         moddat_g_int_CntErr = 0
      End If
   End If

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT "
   g_str_Parame = g_str_Parame & " CASE WHEN  SUBSTR(CNTA_CTBL,1,2) =41 THEN"
   g_str_Parame = g_str_Parame & "       'Gastos Financieros'"
   g_str_Parame = g_str_Parame & "    WHEN  LPAD(CNTA_CTBL,2) =42THEN"
   g_str_Parame = g_str_Parame & "       'Gastos por servicioss'"
   g_str_Parame = g_str_Parame & "    WHEN  LPAD(CNTA_CTBL,2) =51 THEN"
   g_str_Parame = g_str_Parame & "       'Ingresos Financieros'"
   g_str_Parame = g_str_Parame & "    WHEN  LPAD(CNTA_CTBL,2) =52 THEN"
   g_str_Parame = g_str_Parame & "       'Ingresos por Servicios'"
   g_str_Parame = g_str_Parame & "   END AS INDBAS"
   g_str_Parame = g_str_Parame & " ,CNTA_CTBL"
   g_str_Parame = g_str_Parame & " , DESC_CNTA"
   g_str_Parame = g_str_Parame & " ,CASE WHEN SUBSTR(CNTA_CTBL,1,2)  IN ('41','42') THEN"
   g_str_Parame = g_str_Parame & "     TO_NUMBER('-'|| SDO_ACT_DEBE)"
   g_str_Parame = g_str_Parame & "    WHEN SUBSTR(CNTA_CTBL,1,2) IN ('51','52') THEN"
   g_str_Parame = g_str_Parame & "      SDO_ACT_HABER"
   g_str_Parame = g_str_Parame & "   END AS IMPORTE"
   g_str_Parame = g_str_Parame & " FROM("
   g_str_Parame = g_str_Parame & "   SELECT *"
   g_str_Parame = g_str_Parame & "   FROM TT_BALANCE_COMPROB_PRUEBA"
   g_str_Parame = g_str_Parame & "   ORDER BY CNTA_CTBL"
   g_str_Parame = g_str_Parame & " )"
   g_str_Parame = g_str_Parame & " WHERE SUBSTR(CNTA_CTBL,1,2) IN ('41','42','51','52')"
   g_str_Parame = g_str_Parame & " ORDER BY CNTA_CTBL"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   ''********************************************
   Set r_obj_Excel = New Excel.Application

   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   'r_obj_Excel.Workbooks.Item.Name ("Totales")


''***********************************************
'  HOJA DE EXCEL 2 DE TOTALES
''***********************************************

   r_int_NroFil = 11
   r_int_NroIni = 1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT INDBAS,CNTA_CTBL, SUM(IMPORTE) AS IMPORTE"
   g_str_Parame = g_str_Parame & " FROM "
   g_str_Parame = g_str_Parame & " ("
   g_str_Parame = g_str_Parame & "   SELECT "
   g_str_Parame = g_str_Parame & "   CASE WHEN  SUBSTR(CNTA_CTBL,1,2) =41 THEN"
   g_str_Parame = g_str_Parame & "         'Gastos Financieros'"
   g_str_Parame = g_str_Parame & "      WHEN  LPAD(CNTA_CTBL,2) =42THEN"
   g_str_Parame = g_str_Parame & "         'Gastos por servicioss'"
   g_str_Parame = g_str_Parame & "      WHEN  LPAD(CNTA_CTBL,2) =51 THEN"
   g_str_Parame = g_str_Parame & "         'Ingresos Financieros'"
   g_str_Parame = g_str_Parame & "      WHEN  LPAD(CNTA_CTBL,2) =52 THEN"
   g_str_Parame = g_str_Parame & "         'Ingresos por Servicios'"
   g_str_Parame = g_str_Parame & "     END AS INDBAS"
   g_str_Parame = g_str_Parame & "   ,CASE WHEN  SUBSTR(CNTA_CTBL,1,2) =41 THEN"
   g_str_Parame = g_str_Parame & "         '4100'"
   g_str_Parame = g_str_Parame & "      WHEN  LPAD(CNTA_CTBL,2) =42THEN"
   g_str_Parame = g_str_Parame & "         '4200'"
   g_str_Parame = g_str_Parame & "      WHEN  LPAD(CNTA_CTBL,2) =51 THEN"
   g_str_Parame = g_str_Parame & "         '5100'"
   g_str_Parame = g_str_Parame & "      WHEN  LPAD(CNTA_CTBL,2) =52 THEN"
   g_str_Parame = g_str_Parame & "         '5200'"
   g_str_Parame = g_str_Parame & "     END AS CNTA_CTBL"
   g_str_Parame = g_str_Parame & "   ,CASE WHEN SUBSTR(CNTA_CTBL,1,2)  IN ('41','42') THEN"
   g_str_Parame = g_str_Parame & "       TO_NUMBER('-'|| SDO_ACT_DEBE)"
   g_str_Parame = g_str_Parame & "      WHEN SUBSTR(CNTA_CTBL,1,2) IN ('51','52') THEN"
   g_str_Parame = g_str_Parame & "        SDO_ACT_HABER"
   g_str_Parame = g_str_Parame & "     END AS IMPORTE"
   g_str_Parame = g_str_Parame & "   FROM("
   g_str_Parame = g_str_Parame & "     SELECT *"
   g_str_Parame = g_str_Parame & "     FROM TT_BALANCE_COMPROB_PRUEBA"
   g_str_Parame = g_str_Parame & "     ORDER BY CNTA_CTBL"
   g_str_Parame = g_str_Parame & "   )"
   g_str_Parame = g_str_Parame & "   WHERE SUBSTR(CNTA_CTBL,1,2) IN ('41','42','51','52')"
   g_str_Parame = g_str_Parame & "   ORDER BY CNTA_CTBL"
   g_str_Parame = g_str_Parame & " )"
   g_str_Parame = g_str_Parame & " GROUP BY INDBAS,CNTA_CTBL"
   g_str_Parame = g_str_Parame & " ORDER BY CNTA_CTBL"
   
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If

   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
      MsgBox "No se encontraron datos registrados.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   
   With r_obj_Excel.ActiveSheet
   
      'IMAGEN
      .Pictures.Insert ("\\Server_micasita\APLICACIONES\Imagenes\image001.jpg")
      .DrawingObjects(1).Left = 20
      .DrawingObjects(1).Top = 0

      'Unir celdas
      .Range("B6") = "Método Indicador Básico - Totales"
      .Range("A6:D6").Merge
      .Range("A6:D6").Font.Underline = True

      .Range("A6:D6").HorizontalAlignment = xlHAlignCenter
      .Range("D1") = "Dpto. Tecnología e Informática"
      .Range("D2") = "Desarrollo de Sistemas"
      .Range("D1").HorizontalAlignment = xlHAlignRight
      .Range("D2").HorizontalAlignment = xlHAlignRight

      'font
      .Range("A1:D20").Font.Name = "Arial"
      .Range("A1:D20").Font.Size = 9

      'WIDTH
      .Columns("A").ColumnWidth = 12
      .Columns("B").ColumnWidth = 20
      .Columns("C").ColumnWidth = 14
      .Columns("D").ColumnWidth = 20

      .Cells(8, 1) = "F. de Inicio : " & ipp_FecIni.Year & _
      "/" & Trim(IIf(Len(Trim(ipp_FecIni.Month)) = 1, 0 & ipp_FecIni.Month, ipp_FecIni.Month)) & _
      "/" & Trim(IIf(Len(Trim(ipp_FecIni.Day)) = 1, 0 & ipp_FecIni.Day, ipp_FecIni.Day))

      .Cells(9, 1) = "F. de Fin     : " & ipp_FecFin.Year & _
      "/" & Trim(IIf(Len(Trim(ipp_FecFin.Month)) = 1, 0 & ipp_FecFin.Month, ipp_FecFin.Month)) & _
      "/" & Trim(IIf(Len(Trim(ipp_FecFin.Day)) = 1, 0 & ipp_FecFin.Day, ipp_FecFin.Day))

      .Range("A4:B4").Merge

      'CABECERA
      .Cells(r_int_NroFil, r_int_ColumA) = "Item"
      .Cells(r_int_NroFil, r_int_ColumB) = "Indicador Básico"
      .Cells(r_int_NroFil, r_int_ColumC) = "Cuenta Contable"
      .Cells(r_int_NroFil, r_int_ColumD) = "Importe"

      .Columns("c").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_NroFil & ":D" & r_int_NroFil & "").Font.Bold = True
      .Range("A" & r_int_NroFil & ":D" & r_int_NroFil & "").Interior.Color = RGB(146, 208, 80)


      'Bordes de las celdas
      .Range("A" & r_int_NroFil & ":D" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("A" & r_int_NroFil & ":D" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeTop).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumB).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeRight).Weight = xlThin
    
      'LISTADO DE INDICADOR BASICO
      g_rst_Princi.MoveFirst

      Do While Not g_rst_Listas.EOF

         r_int_NroFil = r_int_NroFil + 1
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumA) = r_int_NroIni
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumB) = Trim(g_rst_Listas!INDBAS)
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumC) = "'" & Trim(g_rst_Listas!CNTA_CTBL)
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumD) = g_rst_Listas!IMPORTE
         r_lng_Import = r_lng_Import + g_rst_Listas!IMPORTE

         r_int_NroIni = r_int_NroIni + 1
         
         'Bordes de las celdas
         .Range("A" & r_int_NroFil & ":D" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeTop).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumB).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeRight).Weight = xlThin

      g_rst_Listas.MoveNext
      Loop
      
      r_int_NroFil = r_int_NroFil + 1
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumB) = "Total"
      .Cells(r_int_NroFil, r_int_ColumB).Font.Bold = True
      
      'Bordes de las celdas
      .Range("B" & r_int_NroFil & ":D" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeTop).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumB).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeRight).Weight = xlThin
      
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumD) = r_lng_Import
      
   End With
   

''***********************************************
'  HOJA DE EXCEL 1
''***********************************************
   r_obj_Excel.Sheets.Add().Name = "Balance_comprobacion"
   
   r_int_NroFil = 11
   r_int_NroIni = 1
   
   With r_obj_Excel.ActiveSheet
      'IMAGEN
      .Pictures.Insert ("\\Server_micasita\APLICACIONES\Imagenes\image001.jpg")
      .DrawingObjects(1).Left = 20
      .DrawingObjects(1).Top = 0
  

      'Unir celdas
      .Range("B6") = "Método Indicador Básico - Detalle"
      .Range("A6:E6").Merge
      .Range("A6:E6").Font.Underline = True

      .Range("A6:E6").HorizontalAlignment = xlHAlignCenter
      .Range("E1") = "Dpto. Tecnología e Informática"
      .Range("E2") = "Desarrollo de Sistemas"
      .Range("E1").HorizontalAlignment = xlHAlignRight
      .Range("E2").HorizontalAlignment = xlHAlignRight

      'font
      .Range("A1:E10").Font.Name = "Arial"
      .Range("A1:E10").Font.Size = 9

      'WIDTH
      .Columns("A").ColumnWidth = 7
      .Columns("B").ColumnWidth = 19.6
      .Columns("C").ColumnWidth = 16.6
      .Columns("D").ColumnWidth = 47
      .Columns("E").ColumnWidth = 19

      .Cells(8, 1) = "F. de Inicio : " & ipp_FecIni.Year & _
      "/" & Trim(IIf(Len(Trim(ipp_FecIni.Month)) = 1, 0 & ipp_FecIni.Month, ipp_FecIni.Month)) & _
      "/" & Trim(IIf(Len(Trim(ipp_FecIni.Day)) = 1, 0 & ipp_FecIni.Day, ipp_FecIni.Day))

      .Cells(9, 1) = "F. de Fin     : " & ipp_FecFin.Year & _
      "/" & Trim(IIf(Len(Trim(ipp_FecFin.Month)) = 1, 0 & ipp_FecFin.Month, ipp_FecFin.Month)) & _
      "/" & Trim(IIf(Len(Trim(ipp_FecFin.Day)) = 1, 0 & ipp_FecFin.Day, ipp_FecFin.Day))

      .Range("A4:B4").Merge

      'CABECERA
      .Cells(r_int_NroFil, r_int_ColumA) = "Item"
      .Cells(r_int_NroFil, r_int_ColumB) = "Indicador Básico"
      .Cells(r_int_NroFil, r_int_ColumC) = "Cuenta Contable"
      .Cells(r_int_NroFil, r_int_ColumD) = "Descripcion"
      .Cells(r_int_NroFil, r_int_ColumE) = "Importe"

      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_NroFil & ":E" & r_int_NroFil & "").Font.Bold = True
      .Range("A" & r_int_NroFil & ":E" & r_int_NroFil & "").Interior.Color = RGB(146, 208, 80)

      'Bordes de las celdas
      .Range("A" & r_int_NroFil & ":E" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("A" & r_int_NroFil & ":E" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeTop).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumB).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumE).Borders(xlEdgeRight).Weight = xlThin
      
      'LISTADO DE INDICADOR BASICO
      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF

         r_int_NroFil = r_int_NroFil + 1
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumA) = r_int_NroIni
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumB) = Trim(g_rst_Princi!INDBAS)
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumC) = "'" & Trim(g_rst_Princi!CNTA_CTBL)
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumD) = Trim(g_rst_Princi!DESC_CNTA)
         r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumE) = Trim(g_rst_Princi!IMPORTE)
         
         'Bordes de las celdas
         .Range("A" & r_int_NroFil & ":E" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeTop).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumA).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumB).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumE).Borders(xlEdgeRight).Weight = xlThin
         
         r_int_NroIni = r_int_NroIni + 1
         
      g_rst_Princi.MoveNext
      Loop

      r_int_NroIni = r_int_NroIni + 3
      'font
      .Range("A11:E" & r_int_NroFil & "").Font.Name = "Arial"
      .Range("A11:E" & r_int_NroFil & "").Font.Size = 10

   End With


'**************

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

   
   r_obj_Excel.Visible = True
   
   ''********************************************

End Sub

Private Sub Form_Load()
   ipp_FecIni.Text = date
   ipp_FecFin.Text = date
End Sub
