VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptSun_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   10365
   ClientTop       =   2835
   ClientWidth     =   5940
   Icon            =   "GesCtb_frm_845.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2475
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5955
      _Version        =   65536
      _ExtentX        =   10504
      _ExtentY        =   4366
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
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
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
            TabIndex        =   8
            Top             =   315
            Width           =   4305
            _Version        =   65536
            _ExtentX        =   7594
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F14.1. - Registro de Ventas e Ingresos"
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
            Picture         =   "GesCtb_frm_845.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   915
         Left            =   30
         TabIndex        =   10
         Top             =   1470
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
         _ExtentY        =   1614
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
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   2500
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1080
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
         Begin VB.Label Label3 
            Caption         =   "Año :"
            Height          =   285
            Left            =   210
            TabIndex        =   12
            Top             =   540
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo :"
            Height          =   315
            Left            =   210
            TabIndex        =   11
            Top             =   180
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   780
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
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
         Begin VB.CommandButton cmd_Archivo 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_845.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Generar archivo texto"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5190
            Picture         =   "GesCtb_frm_845.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   60
            Picture         =   "GesCtb_frm_845.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_RptSun_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_bol_FlgReg     As Boolean
Dim r_int_PerMes     As Integer
Dim r_int_PerAno     As Integer

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
   
   'Verifica que exista ruta
   If Dir$(moddat_g_str_RutLoc, vbDirectory) = "" Then
      MsgBox "Debe crear el siguente directorio " & moddat_g_str_RutLoc, vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   r_int_PerMes = CInt(cmb_CodMes.ItemData(cmb_CodMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Archivo_Click()
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

   Screen.MousePointer = 11
   r_int_PerMes = CInt(cmb_CodMes.ItemData(cmb_CodMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   Call fs_GenExpArc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_CodMes, 1, "033")
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_CodMes)
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
Dim r_rst_Total      As ADODB.Recordset
Dim r_obj_Excel      As Excel.Application
Dim r_int_CntFil     As Integer
Dim r_int_NroCor     As Integer
Dim r_str_NumRuc     As String
Dim r_str_NumFac     As String

   r_int_NroCor = 1
   r_int_CntFil = 13

   '********
   'Facturas
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.NRO_LIBRO, A.NRO_ASIENTO, A.ITEM, A.CNTA_CTBL, A.FECHA_CNTBL, A.DET_GLOSA, A.FLAG_DEBHAB,"
   g_str_Parame = g_str_Parame & "       DECODE(FLAG_DEBHAB,'D',-A.IMP_MOVSOL,A.IMP_MOVSOL) AS IMP_MOVSOL, B.TASA_CAMBIO"
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO_DET A"
   g_str_Parame = g_str_Parame & " INNER JOIN CNTBL_ASIENTO B ON B.ORIGEN = A.ORIGEN AND B.ANO = A.ANO AND B.MES = A.MES AND B.NRO_LIBRO = A.NRO_LIBRO AND B.NRO_ASIENTO = A.NRO_ASIENTO"
   g_str_Parame = g_str_Parame & " WHERE (A.ANO = " & r_int_PerAno & ""
   g_str_Parame = g_str_Parame & "   AND A.MES = " & r_int_PerMes & ")"
   g_str_Parame = g_str_Parame & "   AND ((A.FLAG_DEBHAB = 'H'"
   g_str_Parame = g_str_Parame & "   AND A.CNTA_CTBL IN ('511504090101','512504090101','521229010114','521229010116','522229010116','561901010101','562901010101','571101010101'))"
   g_str_Parame = g_str_Parame & "    OR (A.NRO_LIBRO = 16"
   g_str_Parame = g_str_Parame & "   AND (CNTA_CTBL IN ('522229010116','521229010116' ,'561901010101','562901010101'))"
   g_str_Parame = g_str_Parame & "   AND FLAG_DEBHAB = 'D'))"
   g_str_Parame = g_str_Parame & " ORDER BY FLAG_DEBHAB DESC,SUBSTR(A.DET_GLOSA,13,4) ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No hay datos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "REGISTRO DE VENTAS E INGRESOS"

   With r_obj_Excel.Sheets(1)
      .Cells(2, 1) = "FORMATO 14.1: REGISTRO DE VENTAS E INGRESOS"
      .Cells(4, 1) = "PERIODO:  " & Trim(cmb_CodMes.Text) & "  " & Trim(ipp_PerAno.Text)
      .Cells(5, 1) = "RUC    :  20511904162"
      .Cells(6, 1) = "RAZON SOCIAL:  EDPYME MICASITA S.A."
      .Cells(7, 1) = "OFICINA:  PRINCIPAL"
      .Rows(12).RowHeight = 30
      
      .Range(.Cells(10, 1), .Cells(12, 1)).Merge
      .Range(.Cells(10, 1), .Cells(12, 1)).WrapText = True
      .Range(.Cells(10, 1), .Cells(12, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 1), .Cells(12, 1)).ColumnWidth = 10
      .Range(.Cells(10, 2), .Cells(12, 2)).Merge
      .Range(.Cells(10, 2), .Cells(12, 2)).WrapText = True
      .Range(.Cells(10, 2), .Cells(12, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 2), .Cells(12, 2)).ColumnWidth = 13.71
      .Range(.Cells(10, 3), .Cells(12, 3)).Merge
      .Range(.Cells(10, 3), .Cells(12, 3)).WrapText = True
      .Range(.Cells(10, 3), .Cells(12, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 3), .Cells(12, 3)).ColumnWidth = 10
      .Range(.Cells(10, 4), .Cells(12, 4)).Merge
      .Range(.Cells(10, 4), .Cells(12, 4)).WrapText = True
      .Range(.Cells(10, 4), .Cells(12, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 4), .Cells(12, 4)).ColumnWidth = 11
      .Range(.Cells(10, 5), .Cells(12, 5)).Merge
      .Range(.Cells(10, 5), .Cells(12, 5)).WrapText = True
      .Range(.Cells(10, 5), .Cells(12, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 5), .Cells(12, 5)).ColumnWidth = 13
      
      'concatenar
      .Range(.Cells(10, 6), .Cells(10, 9)).Merge
      .Range(.Cells(10, 6), .Cells(10, 9)).WrapText = True
      .Range(.Cells(10, 6), .Cells(10, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 6), .Cells(12, 6)).Merge
      .Range(.Cells(11, 6), .Cells(12, 6)).WrapText = True
      .Range(.Cells(11, 6), .Cells(12, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 6), .Cells(12, 6)).ColumnWidth = 10
      .Range(.Cells(11, 7), .Cells(12, 7)).Merge
      .Range(.Cells(11, 7), .Cells(12, 7)).WrapText = True
      .Range(.Cells(11, 7), .Cells(12, 7)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 7), .Cells(12, 7)).ColumnWidth = 11
      .Range(.Cells(11, 8), .Cells(12, 8)).Merge
      .Range(.Cells(11, 8), .Cells(12, 8)).WrapText = True
      .Range(.Cells(11, 8), .Cells(12, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 8), .Cells(12, 8)).ColumnWidth = 12
      .Range(.Cells(11, 9), .Cells(12, 9)).Merge
      .Range(.Cells(11, 9), .Cells(12, 9)).WrapText = True
      .Range(.Cells(11, 9), .Cells(12, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 9), .Cells(12, 9)).ColumnWidth = 12
      
      'concatenar
      .Range(.Cells(10, 10), .Cells(10, 12)).Merge
      .Range(.Cells(10, 10), .Cells(10, 12)).WrapText = True
      .Range(.Cells(10, 10), .Cells(10, 12)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 10
      .Cells(12, 10).WrapText = True
      .Cells(12, 10).HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 12
      .Cells(12, 11).WrapText = True
      .Cells(12, 11).HorizontalAlignment = xlHAlignCenter
      
      'concatenar
      .Range(.Cells(11, 10), .Cells(11, 11)).Merge
      .Range(.Cells(11, 10), .Cells(11, 11)).WrapText = True
      .Range(.Cells(11, 10), .Cells(11, 11)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 12), .Cells(12, 12)).Merge
      .Range(.Cells(11, 12), .Cells(12, 12)).WrapText = True
      .Range(.Cells(11, 12), .Cells(12, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 12), .Cells(12, 12)).ColumnWidth = 41
      .Range(.Cells(10, 13), .Cells(12, 13)).Merge
      .Range(.Cells(10, 13), .Cells(12, 13)).WrapText = True
      .Range(.Cells(10, 13), .Cells(12, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 13), .Cells(12, 13)).ColumnWidth = 15
      .Range(.Cells(10, 14), .Cells(12, 14)).Merge
      .Range(.Cells(10, 14), .Cells(12, 14)).WrapText = True
      .Range(.Cells(10, 14), .Cells(12, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 14), .Cells(12, 14)).ColumnWidth = 16
      
      'concatenar
      .Range(.Cells(10, 15), .Cells(10, 16)).Merge
      .Range(.Cells(10, 15), .Cells(10, 16)).WrapText = True
      .Range(.Cells(10, 15), .Cells(10, 16)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 15), .Cells(12, 15)).Merge
      .Range(.Cells(11, 15), .Cells(12, 15)).WrapText = True
      .Range(.Cells(11, 15), .Cells(12, 15)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 15), .Cells(12, 15)).ColumnWidth = 14
      .Range(.Cells(11, 16), .Cells(12, 16)).Merge
      .Range(.Cells(11, 16), .Cells(12, 16)).WrapText = True
      .Range(.Cells(11, 16), .Cells(12, 16)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 16), .Cells(12, 16)).ColumnWidth = 16
      .Range(.Cells(10, 17), .Cells(12, 17)).Merge
      .Range(.Cells(10, 17), .Cells(12, 17)).WrapText = True
      .Range(.Cells(10, 17), .Cells(12, 17)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 17), .Cells(12, 17)).ColumnWidth = 12
      .Range(.Cells(10, 18), .Cells(12, 18)).Merge
      .Range(.Cells(10, 18), .Cells(12, 18)).WrapText = True
      .Range(.Cells(10, 18), .Cells(12, 18)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 18), .Cells(12, 18)).ColumnWidth = 12
      .Range(.Cells(10, 19), .Cells(12, 19)).Merge
      .Range(.Cells(10, 19), .Cells(12, 19)).WrapText = True
      .Range(.Cells(10, 19), .Cells(12, 19)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 19), .Cells(12, 19)).ColumnWidth = 13
      .Range(.Cells(10, 20), .Cells(12, 20)).Merge
      .Range(.Cells(10, 20), .Cells(12, 20)).WrapText = True
      .Range(.Cells(10, 20), .Cells(12, 20)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 20), .Cells(12, 20)).ColumnWidth = 13
      .Range(.Cells(10, 21), .Cells(12, 21)).Merge
      .Range(.Cells(10, 21), .Cells(12, 21)).WrapText = True
      .Range(.Cells(10, 21), .Cells(12, 21)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 21), .Cells(12, 21)).ColumnWidth = 18
      .Range(.Cells(10, 22), .Cells(12, 22)).Merge
      .Range(.Cells(10, 22), .Cells(12, 22)).WrapText = True
      .Range(.Cells(10, 22), .Cells(12, 22)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 22), .Cells(12, 22)).ColumnWidth = 18
      .Range(.Cells(10, 23), .Cells(12, 23)).Merge
      .Range(.Cells(10, 23), .Cells(12, 23)).WrapText = True
      .Range(.Cells(10, 23), .Cells(12, 23)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 23), .Cells(12, 23)).ColumnWidth = 8
      
      'concatenar
      .Range(.Cells(10, 24), .Cells(10, 27)).Merge
      .Range(.Cells(10, 24), .Cells(10, 27)).WrapText = True
      .Range(.Cells(10, 24), .Cells(10, 27)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 24), .Cells(12, 24)).Merge
      .Range(.Cells(11, 24), .Cells(12, 24)).WrapText = True
      .Range(.Cells(11, 24), .Cells(12, 24)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 24), .Cells(12, 24)).ColumnWidth = 10
      .Range(.Cells(11, 25), .Cells(12, 25)).Merge
      .Range(.Cells(11, 25), .Cells(12, 25)).WrapText = True
      .Range(.Cells(11, 25), .Cells(12, 25)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 25), .Cells(12, 25)).ColumnWidth = 10
      .Range(.Cells(11, 26), .Cells(12, 26)).Merge
      .Range(.Cells(11, 26), .Cells(12, 26)).WrapText = True
      .Range(.Cells(11, 26), .Cells(12, 26)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 26), .Cells(12, 26)).ColumnWidth = 10
      .Range(.Cells(11, 27), .Cells(12, 27)).Merge
      .Range(.Cells(11, 27), .Cells(12, 27)).WrapText = True
      .Range(.Cells(11, 27), .Cells(12, 27)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 27), .Cells(12, 27)).ColumnWidth = 15
      .Range(.Cells(10, 28), .Cells(12, 28)).Merge
      .Range(.Cells(10, 28), .Cells(12, 28)).WrapText = True
      .Range(.Cells(10, 28), .Cells(12, 28)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 28), .Cells(12, 28)).ColumnWidth = 13
      .Range(.Cells(10, 29), .Cells(12, 29)).Merge
      .Range(.Cells(10, 29), .Cells(12, 29)).WrapText = True
      .Range(.Cells(10, 29), .Cells(12, 29)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 29), .Cells(12, 29)).ColumnWidth = 10
      
      .Cells(10, 1) = "PERIODO"
      .Cells(10, 2) = "NUMERO CORRELATIVO"
      .Cells(10, 3) = "NUMERO CUO"
      .Cells(10, 4) = "FECHA DE EMISION"
      .Cells(10, 5) = "FECHA DE VENCIMIENTO Y/O PAGO"
      .Cells(10, 6) = "COMPROBANTE DE PAGO O DOCUMENTO"
      .Cells(11, 6) = "TIPO (TABLA 10)"
      .Cells(11, 7) = "NRO. SERIE"
      .Cells(11, 8) = "NUMERO INICIAL"
      .Cells(11, 9) = "NUMERO FINAL"
      .Cells(10, 10) = "INFORMACION DEL CLIENTE"
      .Cells(11, 10) = "DOCUMENTO IDENTIDAD"
      .Cells(12, 10) = "TIPO (TABLA 2)"
      .Cells(12, 11) = "NUMERO"
      .Cells(11, 12) = "APELLIDOS Y NOMBRES DENOMINACION O RAZON SOCIAL"
      .Cells(10, 13) = "VALOR FACTURADO DE EXPORTACION"
      .Cells(10, 14) = "BASE IMPONIBLE (BI) OPERACION GRAVADA"
      .Cells(10, 15) = "IMPORTE TOTAL DE LA OPERACION"
      .Cells(11, 15) = "EXONERADA"
      .Cells(11, 16) = "INAFECTA"
      .Cells(10, 17) = "ISC"
      .Cells(10, 18) = "IGV Y/O IPM"
      .Cells(10, 19) = "BI OPERACION GRAVADA CON IMPUES.VTAS."
      .Cells(10, 20) = "IMPUESTO A LAS VENTAS DEL ARROZ PILADO"
      .Cells(10, 21) = "OTROS TRIBUTOS Y CARGOS QUE NO FORMAN PARTE DE BASE IMPONIBLE"
      .Cells(10, 22) = "IMPORTE TOTAL DEL COMPROBANTE DE PAGO (EQUIV. MN)"
      .Cells(10, 23) = "TIPO DE CAMBIO"
      .Cells(10, 24) = "REFERENCIA DEL COMPROBANTE DE PAGO O DOCUMENTO "
      .Cells(11, 24) = "FECHA"
      .Cells(11, 25) = "TIPO (TABLA 10)"
      .Cells(11, 26) = "SERIE"
      .Cells(11, 27) = "NRO.COMPROB. DE PAGO O DOCUMENTO"
      .Cells(10, 28) = "VALOR EMBARCADO DE EXPORTACION"
      .Cells(10, 29) = "ESTADO"
      
      .Range(.Cells(1, 1), .Cells(12, 29)).Font.Bold = True
      .Range(.Cells(10, 1), .Cells(12, 29)).Interior.Color = RGB(146, 208, 80)
      
      .Columns("M").NumberFormat = "#,###,##0.00"
      .Columns("N").NumberFormat = "#,###,##0.00"
      .Columns("O").NumberFormat = "#,###,##0.00"
      .Columns("P").NumberFormat = "#,###,##0.00"
      .Columns("Q").NumberFormat = "#,###,##0.00"
      .Columns("R").NumberFormat = "#,###,##0.00"
      .Columns("S").NumberFormat = "#,###,##0.00"
      .Columns("T").NumberFormat = "#,###,##0.00"
      .Columns("U").NumberFormat = "#,###,##0.00"
      .Columns("V").NumberFormat = "#,###,##0.00"
      .Columns("W").NumberFormat = "###,##0.0000"
      .Columns("H").NumberFormat = "@"
      
      .Range(.Cells(10, 1), .Cells(12, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      Do Until r_int_NroCor > 29
         .Range(.Cells(12, r_int_NroCor), .Cells(12, r_int_NroCor)).VerticalAlignment = xlCenter
         .Range(.Cells(10, r_int_NroCor), .Cells(10, r_int_NroCor)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(12, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(10, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(10, r_int_NroCor), .Cells(10, r_int_NroCor)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(12, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(10, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeRight).LineStyle = xlContinuous
         r_int_NroCor = r_int_NroCor + 1
      Loop
      
      .Range(.Cells(10, 6), .Cells(10, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(11, 10), .Cells(11, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(10, 10), .Cells(10, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(10, 13), .Cells(10, 29)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroCor = 1
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If InStr(Trim(g_rst_Princi!DET_GLOSA), "B/V") > 0 Then
            r_str_NumRuc = ""
            r_str_NumFac = Trim(Mid(Trim(g_rst_Princi!DET_GLOSA), 4, 10))
         Else
            r_str_NumRuc = Trim(Mid(Trim(g_rst_Princi!DET_GLOSA), 1, 11))
            r_str_NumFac = Trim(Mid(Trim(g_rst_Princi!DET_GLOSA), 13, 3))
         End If
         If InStr(Trim(g_rst_Princi!DET_GLOSA), "B/V") > 0 Then
            GoTo Ingresar
         Else
            If gf_Valida_RUC(r_str_NumRuc, Mid(r_str_NumRuc, 11, 1)) Then
Ingresar:
               .Cells(r_int_CntFil, 1) = Format(g_rst_Princi!FECHA_CNTBL, "YYYYMM00")
               .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 1)).HorizontalAlignment = xlHAlignCenter
               .Cells(r_int_CntFil, 2) = r_int_NroCor
               .Range(.Cells(r_int_CntFil, 2), .Cells(r_int_CntFil, 2)).HorizontalAlignment = xlHAlignCenter
               .Cells(r_int_CntFil, 3) = "M" & r_int_NroCor
               .Range(.Cells(r_int_CntFil, 3), .Cells(r_int_CntFil, 3)).HorizontalAlignment = xlHAlignCenter
               .Cells(r_int_CntFil, 4) = "'" & Trim(g_rst_Princi!FECHA_CNTBL)
               .Range(.Cells(r_int_CntFil, 4), .Cells(r_int_CntFil, 4)).HorizontalAlignment = xlHAlignCenter
               .Cells(r_int_CntFil, 5) = ""
               .Cells(r_int_CntFil, 6) = "'" & "01"
               .Range(.Cells(r_int_CntFil, 6), .Cells(r_int_CntFil, 6)).HorizontalAlignment = xlHAlignCenter
               .Cells(r_int_CntFil, 7) = "'" & "001"
               .Range(.Cells(r_int_CntFil, 7), .Cells(r_int_CntFil, 7)).HorizontalAlignment = xlHAlignCenter
               .Cells(r_int_CntFil, 8) = Trim(r_str_NumFac)
               .Range(.Cells(r_int_CntFil, 8), .Cells(r_int_CntFil, 8)).HorizontalAlignment = xlHAlignCenter
               .Cells(r_int_CntFil, 9) = ""
               .Cells(r_int_CntFil, 10) = "6"
               .Range(.Cells(r_int_CntFil, 10), .Cells(r_int_CntFil, 10)).HorizontalAlignment = xlHAlignCenter
               .Cells(r_int_CntFil, 11) = r_str_NumRuc
               .Range(.Cells(r_int_CntFil, 11), .Cells(r_int_CntFil, 11)).HorizontalAlignment = xlHAlignCenter
               
               If InStr(Trim(g_rst_Princi!DET_GLOSA), "B/V") > 0 Then
                  .Cells(r_int_CntFil, 12) = Trim(Mid(Trim(g_rst_Princi!DET_GLOSA), 13))
               Else
                  .Cells(r_int_CntFil, 12) = Trim(Mid(Trim(g_rst_Princi!DET_GLOSA), 16, Len(Trim(g_rst_Princi!DET_GLOSA)) - 17))
               End If
               .Cells(r_int_CntFil, 13) = 0
               .Cells(r_int_CntFil, 14) = Trim(g_rst_Princi!IMP_MOVSOL)
   '            If InStr(r_str_NumFac, "NC") > 0 Then r_dbl_TotNCr = r_dbl_TotNCr + Trim(g_rst_Princi!IMP_MOVSOL)
               .Cells(r_int_CntFil, 15) = 0
               .Cells(r_int_CntFil, 16) = 0
               .Cells(r_int_CntFil, 17) = 0
               .Cells(r_int_CntFil, 18) = g_rst_Princi!IMP_MOVSOL * 0.18
               .Cells(r_int_CntFil, 19) = 0
               .Cells(r_int_CntFil, 20) = 0
               .Cells(r_int_CntFil, 21) = 0
               .Cells(r_int_CntFil, 22) = g_rst_Princi!IMP_MOVSOL * 1.18
               If Mid(Trim(g_rst_Princi!CNTA_CTBL), 3, 1) = "2" Then
                  .Cells(r_int_CntFil, 23) = g_rst_Princi!TASA_CAMBIO
               Else
                  .Cells(r_int_CntFil, 23) = 0
               End If
               .Cells(r_int_CntFil, 24) = ""
               .Cells(r_int_CntFil, 25) = ""
               .Cells(r_int_CntFil, 26) = ""
               .Cells(r_int_CntFil, 27) = ""
               .Cells(r_int_CntFil, 28) = ""
               .Cells(r_int_CntFil, 29) = "1"
               .Range(.Cells(r_int_CntFil, 29), .Cells(r_int_CntFil, 29)).HorizontalAlignment = xlHAlignCenter
               
               .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeLeft).LineStyle = xlContinuous
               .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeTop).LineStyle = xlContinuous
               .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeBottom).LineStyle = xlContinuous
               .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeRight).LineStyle = xlContinuous
               .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlInsideVertical).LineStyle = xlContinuous
               
               r_int_NroCor = r_int_NroCor + 1
               r_int_CntFil = r_int_CntFil + 1
            End If
         End If
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      
      '**************************
      'Total ingresos financieros
      
      If r_int_PerMes = 12 Then
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "  SELECT SUM(NVL(B.IMP_MOVSOL,0.0)) + NVL(SUM(NVL(A.SLDO_SOLHAB,0) - NVL(A.SLDO_SOLDEB,0)),0) AS SALDO "
         g_str_Parame = g_str_Parame & "    FROM CNTBL_CNTA_SDO A "
         g_str_Parame = g_str_Parame & "         LEFT JOIN CNTBL_ASIENTO_DET B ON B.ORIGEN = 'LM' AND B.ANO = A.ANO AND B.MES = A.MES AND B.NRO_LIBRO = 11 "
         g_str_Parame = g_str_Parame & "                             AND B.NRO_ASIENTO = 2 AND B.CNTA_CTBL = A.CNTA_CTBL"
         g_str_Parame = g_str_Parame & "   WHERE A.ANO = " & r_int_PerAno & " AND A.MES = " & r_int_PerMes & ""
         g_str_Parame = g_str_Parame & "     AND A.CNTA_CTBL IN ( '511401042302','511401042304','511401042305','511401042309','511401042310',"
         g_str_Parame = g_str_Parame & "                          '511405042302','511405042309','511405042310','511405042305','511401042501',"
         g_str_Parame = g_str_Parame & "                          '511401042502','511405042501','511405042502','511401042301','511401042303',"
         g_str_Parame = g_str_Parame & "                          '511405042301','511405042303','511401040601','511401040602','512401040601',"
         g_str_Parame = g_str_Parame & "                          '512401040602','512405040601','512405040602','511405040601','512401042401',"
         g_str_Parame = g_str_Parame & "                          '512401042402','512405042401','512405042402','511401122701','511401132701',"
         g_str_Parame = g_str_Parame & "                          '511704040101','511704040106','521229010109','521229010113','521229010115',"
         g_str_Parame = g_str_Parame & "                          '522229010101','522229010102','522229010109','522229010113','521229010110',"
         g_str_Parame = g_str_Parame & "                          '521229010101','521229010114','511401040603','511401042503','571101010101',"
         g_str_Parame = g_str_Parame & "                          '511926010101')"
         
      Else
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT SUM(SLDO_SOLHAB)-SUM(SLDO_SOLDEB) AS SALDO FROM CNTBL_CNTA_SDO A "
         g_str_Parame = g_str_Parame & " WHERE A.ANO = " & r_int_PerAno & ""
         g_str_Parame = g_str_Parame & "   AND A.MES = " & r_int_PerMes & ""
         g_str_Parame = g_str_Parame & "   AND A.CNTA_CTBL IN ('511401042302','511401042304','511401042305','511401042309','511401042310','511405042302','511405042309','511405042310',"
         g_str_Parame = g_str_Parame & "                       '511405042305','511401042501','511401042502','511405042501','511405042502','511401042301','511401042303','511405042301',"
         g_str_Parame = g_str_Parame & "                       '511405042303','511401040601','511401040602','512401040601','512401040602','512405040601','512405040602','511405040601',"
         g_str_Parame = g_str_Parame & "                       '512401042401','512401042402','512405042401','512405042402','511401122701','511401132701',"
         g_str_Parame = g_str_Parame & "                       '511704040101','511704040106','521229010109','521229010113','521229010115','522229010101','522229010102','522229010109',"
         g_str_Parame = g_str_Parame & "                       '522229010113','521229010110','511504090102','511505010101','512504090102','512505010101','521229010117','522229010117',"
         g_str_Parame = g_str_Parame & "                       '521229010101','511401040603','511401042503','521229010101',"
         g_str_Parame = g_str_Parame & "                       '511926010101', "
         g_str_Parame = g_str_Parame & "                       '561101010101') " 'INGRESO DE NUEVA CUENTA
      End If
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Total, 3) Then
         Exit Sub
      End If
      
      If Not r_rst_Total.EOF And Not r_rst_Total.BOF Then
         .Cells(r_int_CntFil, 1) = r_int_PerAno & Format(r_int_PerMes, "00") & "00"
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 1)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_CntFil, 2) = r_int_NroCor
         .Range(.Cells(r_int_CntFil, 2), .Cells(r_int_CntFil, 2)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_CntFil, 3) = "M" & r_int_NroCor
         .Range(.Cells(r_int_CntFil, 3), .Cells(r_int_CntFil, 3)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_CntFil, 4) = ff_Ultimo_Dia_Mes(r_int_PerMes, Val(r_int_PerAno)) & "/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
         .Range(.Cells(r_int_CntFil, 4), .Cells(r_int_CntFil, 4)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_CntFil, 5) = ""
         .Cells(r_int_CntFil, 6) = "'" & "13"
         .Range(.Cells(r_int_CntFil, 6), .Cells(r_int_CntFil, 6)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_CntFil, 7) = "'"
         .Cells(r_int_CntFil, 8) = ""
         .Cells(r_int_CntFil, 9) = ""
         .Cells(r_int_CntFil, 10) = ""
         .Cells(r_int_CntFil, 11) = ""
         .Cells(r_int_CntFil, 12) = "Consolidado Total Ingresos Financieros"
         .Cells(r_int_CntFil, 13) = 0
         .Cells(r_int_CntFil, 14) = 0
         .Cells(r_int_CntFil, 15) = Trim(r_rst_Total!SALDO)
         .Cells(r_int_CntFil, 16) = 0
         .Cells(r_int_CntFil, 17) = 0
         .Cells(r_int_CntFil, 18) = 0
         .Cells(r_int_CntFil, 19) = 0
         .Cells(r_int_CntFil, 20) = 0
         .Cells(r_int_CntFil, 21) = 0
         .Cells(r_int_CntFil, 22) = Trim(r_rst_Total!SALDO)
         .Cells(r_int_CntFil, 23) = 0
         .Cells(r_int_CntFil, 24) = ""
         .Cells(r_int_CntFil, 25) = ""
         .Cells(r_int_CntFil, 26) = ""
         .Cells(r_int_CntFil, 27) = ""
         .Cells(r_int_CntFil, 28) = ""
         .Cells(r_int_CntFil, 29) = "1"
         .Range(.Cells(r_int_CntFil, 29), .Cells(r_int_CntFil, 29)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlInsideVertical).LineStyle = xlContinuous
         
         r_int_NroCor = r_int_NroCor + 1
         r_int_CntFil = r_int_CntFil + 1
      End If
      
      r_rst_Total.Close
      Set r_rst_Total = Nothing
            
      '****************************************
      'Total ingresos por servicios financieros
      g_str_Parame = ""
      If r_int_PerMes = 12 Then
         g_str_Parame = g_str_Parame & "SELECT SUM(SLDO_SOLHAB)AS SALDO FROM CNTBL_CNTA_SDO "
      Else
         g_str_Parame = g_str_Parame & "SELECT SUM(SLDO_SOLHAB)-SUM(SLDO_SOLDEB) AS SALDO FROM CNTBL_CNTA_SDO "
      End If
      g_str_Parame = g_str_Parame & " WHERE ANO = " & r_int_PerAno & ""
      g_str_Parame = g_str_Parame & "   AND MES = " & r_int_PerMes & ""
      g_str_Parame = g_str_Parame & "   AND CNTA_CTBL IN ('511103010601','511103010901','511103012103','511103012903','511103013201','511103013202','512103010601','512103013201')"

      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Total, 3) Then
         Exit Sub
      End If
      
      If Not r_rst_Total.EOF And Not r_rst_Total.BOF Then
         .Cells(r_int_CntFil, 1) = r_int_PerAno & Format(r_int_PerMes, "00") & "00"
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 1)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_CntFil, 2) = r_int_NroCor
         .Range(.Cells(r_int_CntFil, 2), .Cells(r_int_CntFil, 2)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_CntFil, 3) = "M" & r_int_NroCor
         .Range(.Cells(r_int_CntFil, 3), .Cells(r_int_CntFil, 3)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_CntFil, 4) = ff_Ultimo_Dia_Mes(r_int_PerMes, Val(r_int_PerAno)) & "/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
         .Range(.Cells(r_int_CntFil, 4), .Cells(r_int_CntFil, 4)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_CntFil, 5) = ""
         .Cells(r_int_CntFil, 6) = "'" & "13"
         .Range(.Cells(r_int_CntFil, 6), .Cells(r_int_CntFil, 6)).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_CntFil, 7) = "'" & ""
         .Cells(r_int_CntFil, 8) = ""
         .Cells(r_int_CntFil, 9) = ""
         .Cells(r_int_CntFil, 10) = ""
         .Cells(r_int_CntFil, 11) = ""
         .Cells(r_int_CntFil, 12) = "Consolidado Total Ingresos por Servicios Financieros"
         .Cells(r_int_CntFil, 13) = 0
         .Cells(r_int_CntFil, 14) = 0
         .Cells(r_int_CntFil, 15) = Trim(r_rst_Total!SALDO)
         .Cells(r_int_CntFil, 16) = 0
         .Cells(r_int_CntFil, 17) = 0
         .Cells(r_int_CntFil, 18) = 0
         .Cells(r_int_CntFil, 19) = 0
         .Cells(r_int_CntFil, 20) = 0
         .Cells(r_int_CntFil, 21) = 0
         .Cells(r_int_CntFil, 22) = Trim(r_rst_Total!SALDO)
         .Cells(r_int_CntFil, 23) = 0
         .Cells(r_int_CntFil, 24) = ""
         .Cells(r_int_CntFil, 25) = ""
         .Cells(r_int_CntFil, 26) = ""
         .Cells(r_int_CntFil, 27) = ""
         .Cells(r_int_CntFil, 28) = ""
         .Cells(r_int_CntFil, 29) = "1"
         .Range(.Cells(r_int_CntFil, 29), .Cells(r_int_CntFil, 29)).HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 29)).Borders(xlInsideVertical).LineStyle = xlContinuous
         r_int_CntFil = r_int_CntFil + 1
      End If
      
      r_rst_Total.Close
      Set r_rst_Total = Nothing
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExpArc()
Dim r_int_NroCor     As Integer
Dim r_str_NomRes     As String
Dim r_str_NumRuc     As String
Dim r_str_NumFac     As String
Dim r_str_DetGlo     As String
Dim r_dbl_TipCam     As Double
Dim r_int_NumRes     As Integer
Dim r_rst_Total      As ADODB.Recordset

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.NRO_LIBRO, A.NRO_ASIENTO, A.ITEM, A.CNTA_CTBL, A.FECHA_CNTBL, A.DET_GLOSA, A.FLAG_DEBHAB,"
   g_str_Parame = g_str_Parame & "       DECODE(FLAG_DEBHAB,'D',-A.IMP_MOVSOL,A.IMP_MOVSOL) AS IMP_MOVSOL, B.TASA_CAMBIO"
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO_DET A"
   g_str_Parame = g_str_Parame & " INNER JOIN CNTBL_ASIENTO B ON B.ORIGEN = A.ORIGEN AND B.ANO = A.ANO AND B.MES = A.MES AND B.NRO_LIBRO = A.NRO_LIBRO AND B.NRO_ASIENTO = A.NRO_ASIENTO"
   g_str_Parame = g_str_Parame & " WHERE (A.ANO = " & r_int_PerAno & ""
   g_str_Parame = g_str_Parame & "   AND A.MES = " & r_int_PerMes & ")"
   g_str_Parame = g_str_Parame & "   AND ((A.FLAG_DEBHAB = 'H'"
   g_str_Parame = g_str_Parame & "   AND A.CNTA_CTBL IN ('511504090101','512504090101','521229010114','521229010116','522229010116','561901010101','562901010101','571101010101'))"
   g_str_Parame = g_str_Parame & "    OR (A.NRO_LIBRO = 16"
   g_str_Parame = g_str_Parame & "   AND (CNTA_CTBL IN ('522229010116','521229010116','561901010101','562901010101' ))"
   g_str_Parame = g_str_Parame & "   AND FLAG_DEBHAB = 'D'))"
   g_str_Parame = g_str_Parame & " ORDER BY FLAG_DEBHAB DESC,SUBSTR(A.DET_GLOSA,13,4) ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No hay datos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Creando Archivo
   r_str_NomRes = moddat_g_str_RutLoc & "\LE20511904162" & r_int_PerAno & Format(r_int_PerMes, "00") & "00" & "14" & "0100001111" & ".txt"
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes

   r_int_NroCor = 1

   Do While Not g_rst_Princi.EOF
      
      If InStr(Trim(g_rst_Princi!DET_GLOSA), "B/V") > 0 Then
         r_str_NumRuc = ""
         r_str_NumFac = Trim(Mid(Trim(g_rst_Princi!DET_GLOSA), 4, 10))
         r_str_DetGlo = Trim(Mid(Trim(g_rst_Princi!DET_GLOSA), 13))
      Else
         r_str_NumRuc = Trim(Mid(Trim(g_rst_Princi!DET_GLOSA), 1, 11))
         r_str_NumFac = Trim(Mid(Trim(g_rst_Princi!DET_GLOSA), 13, 3))
         r_str_DetGlo = Trim(Mid(Trim(g_rst_Princi!DET_GLOSA), 16, Len(Trim(g_rst_Princi!DET_GLOSA)) - 17))
      End If
      
      If Mid(Trim(g_rst_Princi!CNTA_CTBL), 3, 1) = "2" Then
         r_dbl_TipCam = g_rst_Princi!TASA_CAMBIO
      Else
         r_dbl_TipCam = 0
      End If
      
      If InStr(Trim(g_rst_Princi!DET_GLOSA), "B/V") > 0 Then
         GoTo Ingresar
      Else

         If gf_Valida_RUC(r_str_NumRuc, Mid(r_str_NumRuc, 11, 1)) Then
Ingresar:
            Print #1, Format(g_rst_Princi!FECHA_CNTBL, "YYYYMM00"); "|"; r_int_NroCor; "|"; "M" & r_int_NroCor; "|"; _
                      Trim(g_rst_Princi!FECHA_CNTBL); "|"; ""; "|"; "01"; "|"; "0001"; "|"; r_str_NumFac; "|"; _
                      "0"; "|"; "6"; "|"; r_str_NumRuc; "|"; Trim(r_str_DetGlo); "|"; "0"; "|"; _
                      Trim(g_rst_Princi!IMP_MOVSOL); "|"; "0"; "|"; "0"; "|"; "0"; "|"; Format(g_rst_Princi!IMP_MOVSOL * 0.18, "#####00.00"); "|"; _
                      "0"; "|"; "0"; "|"; "0"; "|"; Format(g_rst_Princi!IMP_MOVSOL * 1.18, "#####00.00"); "|"; r_dbl_TipCam; "|"; "01/01/0001"; "|"; "00"; "|"; "-"; "|"; _
                      "-"; "|"; "1"; "|"; "1"; "|"; "-"
            
            r_int_NroCor = r_int_NroCor + 1
         End If
      End If
      g_rst_Princi.MoveNext
      DoEvents
   Loop
            
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Total ingresos financieros
   g_str_Parame = ""
   If r_int_PerMes = 12 Then
         g_str_Parame = g_str_Parame & "  SELECT SUM(NVL(B.IMP_MOVSOL,0.0)) + NVL(SUM(NVL(A.SLDO_SOLHAB,0) - NVL(A.SLDO_SOLDEB,0)),0) AS SALDO "
         g_str_Parame = g_str_Parame & "    FROM CNTBL_CNTA_SDO A "
         g_str_Parame = g_str_Parame & "         LEFT JOIN CNTBL_ASIENTO_DET B ON B.ORIGEN = 'LM' AND B.ANO = A.ANO AND B.MES = A.MES AND B.NRO_LIBRO = 11 "
         g_str_Parame = g_str_Parame & "                             AND B.NRO_ASIENTO = 2 AND B.CNTA_CTBL = A.CNTA_CTBL"
         g_str_Parame = g_str_Parame & "   WHERE A.ANO = " & r_int_PerAno & " AND A.MES = " & r_int_PerMes & ""
         g_str_Parame = g_str_Parame & "     AND A.CNTA_CTBL IN ( '511401042302','511401042304','511401042305','511401042309','511401042310',"
         g_str_Parame = g_str_Parame & "                          '511405042302','511405042309','511405042310','511405042305','511401042501',"
         g_str_Parame = g_str_Parame & "                          '511401042502','511405042501','511405042502','511401042301','511401042303',"
         g_str_Parame = g_str_Parame & "                          '511405042301','511405042303','511401040601','511401040602','512401040601',"
         g_str_Parame = g_str_Parame & "                          '512401040602','512405040601','512405040602','511405040601','512401042401',"
         g_str_Parame = g_str_Parame & "                          '512401042402','512405042401','512405042402','511401122701','511401132701',"
         g_str_Parame = g_str_Parame & "                          '511704040101','511704040106','521229010109','521229010113','521229010115',"
         g_str_Parame = g_str_Parame & "                          '522229010101','522229010102','522229010109','522229010113','521229010110',"
         g_str_Parame = g_str_Parame & "                          '521229010101','521229010114','511401040603','511401042503','571101010101',"
         g_str_Parame = g_str_Parame & "                          '511926010101')"
         
   Else
         g_str_Parame = g_str_Parame & "SELECT SUM(SLDO_SOLHAB)-SUM(SLDO_SOLDEB) AS SALDO FROM CNTBL_CNTA_SDO "
         g_str_Parame = g_str_Parame & " WHERE ANO = " & r_int_PerAno & ""
         g_str_Parame = g_str_Parame & "   AND MES = " & r_int_PerMes & ""
         g_str_Parame = g_str_Parame & "   AND CNTA_CTBL IN ('511401042302','511401042304','511401042305','511401042309','511401042310','511405042302','511405042309','511405042310',"
         g_str_Parame = g_str_Parame & "                     '511405042305','511401042501','511401042502','511405042501','511405042502','511401042301','511401042303','511405042301',"
         g_str_Parame = g_str_Parame & "                     '511405042303','511401040601','511401040602','512401040601','512401040602','512405040601','512405040602','511405040601',"
         g_str_Parame = g_str_Parame & "                     '512401042401','512401042402','512405042401','512405042402','511401122701','511401132701',"
         g_str_Parame = g_str_Parame & "                     '511704040101','511704040106','521229010109','521229010113','521229010115','522229010101','522229010102','522229010109',"
         g_str_Parame = g_str_Parame & "                     '522229010113','521229010110','511504090102','511505010101','512504090102','512505010101','521229010117','522229010117',"
         g_str_Parame = g_str_Parame & "                     '521229010101','511401040603','511401042503','521229010101',"
         g_str_Parame = g_str_Parame & "                     '511926010101', "
         g_str_Parame = g_str_Parame & "                     '561101010101')" 'CUENTA FALTANTE
   End If
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Total, 3) Then
      Exit Sub
   End If

   If Not r_rst_Total.EOF And Not r_rst_Total.BOF Then
      Print #1, r_int_PerAno & Format(r_int_PerMes, "00") & "00"; "|"; r_int_NroCor; "|"; "M" & r_int_NroCor; "|"; _
             ff_Ultimo_Dia_Mes(r_int_PerMes, Val(r_int_PerAno)) & "/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno; "|"; _
             ""; "|"; "13"; "|"; "-"; "|"; "1"; "|"; "0"; "|"; ""; "|"; ""; "|"; _
             "Consolidado Total Ingresos Financieros"; "|"; "0"; "|"; "0"; "|"; _
             Format(r_rst_Total!SALDO, "#####00.00"); "|"; "0"; "|"; "0"; "|"; "0"; "|"; _
             "0"; "|"; "0"; "|"; "0"; "|"; Format(r_rst_Total!SALDO, "#####00.00"); "|"; _
             "0.000"; "|"; "01/01/0001"; "|"; "00"; "|"; "-"; "|"; "-"; "|"; "1"; "|"; "1"; "|"; "-"
      
      r_int_NroCor = r_int_NroCor + 1
   End If
   
   r_rst_Total.Close
   Set r_rst_Total = Nothing
   
   'Total ingresos por servicios financieros
   g_str_Parame = ""
   If r_int_PerMes = 12 Then
      g_str_Parame = g_str_Parame & "SELECT SUM(SLDO_SOLHAB)AS SALDO FROM CNTBL_CNTA_SDO "
   Else
      g_str_Parame = g_str_Parame & "SELECT SUM(SLDO_SOLHAB)-SUM(SLDO_SOLDEB) AS SALDO FROM CNTBL_CNTA_SDO "
   End If
   g_str_Parame = g_str_Parame & " WHERE ANO = " & r_int_PerAno & ""
   g_str_Parame = g_str_Parame & "   AND MES = " & r_int_PerMes & ""
   g_str_Parame = g_str_Parame & "   AND CNTA_CTBL IN ('511103010601','511103010901','511103012103','511103012903','511103013201','511103013202','512103010601','512103013201')"

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Total, 3) Then
      Exit Sub
   End If

   If Not r_rst_Total.EOF And Not r_rst_Total.BOF Then
      Print #1, r_int_PerAno & Format(r_int_PerMes, "00") & "00"; "|"; r_int_NroCor; "|"; "M" & r_int_NroCor; "|"; _
          ff_Ultimo_Dia_Mes(r_int_PerMes, Val(r_int_PerAno)) & "/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno; "|"; _
          ""; "|"; "13"; "|"; "-"; "|"; "1"; "|"; "0"; "|"; ""; "|"; ""; "|"; _
          "Consolidado Total Ingresos por Servicios Financieros"; "|"; "0"; "|"; "0"; "|"; _
          Format(r_rst_Total!SALDO, "#####00.00"); "|"; "0"; "|"; "0"; "|"; "0"; "|"; _
          "0"; "|"; "0"; "|"; "0"; "|"; Format(r_rst_Total!SALDO, "#####00.00"); "|"; _
          "0.000"; "|"; "01/01/0001"; "|"; "00"; "|"; "-"; "|"; "-"; "|"; "1"; "|"; "1"; "|"; "-"
   End If
   
   Close #1
   r_rst_Total.Close
   Set r_rst_Total = Nothing
   
   MsgBox "El archivo ha sido creado: " & Trim(r_str_NomRes), vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub cmb_CodMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub
