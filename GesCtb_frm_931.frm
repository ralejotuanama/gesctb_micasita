VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form Frm_Ctb_FacEle_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15885
   Icon            =   "GesCtb_frm_931.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   15885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8800
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15885
      _Version        =   65536
      _ExtentX        =   28019
      _ExtentY        =   15522
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
         TabIndex        =   9
         Top             =   60
         Width           =   15765
         _Version        =   65536
         _ExtentX        =   27808
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
         Begin Threed.SSPanel pnl_Titulo 
            Height          =   300
            Left            =   720
            TabIndex        =   12
            Top             =   210
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Documentos Electrónicos"
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
            Left            =   120
            Picture         =   "GesCtb_frm_931.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6645
         Left            =   60
         TabIndex        =   10
         Top             =   2100
         Width           =   15765
         _Version        =   65536
         _ExtentX        =   27808
         _ExtentY        =   11721
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
            Height          =   6225
            Left            =   30
            TabIndex        =   1
            Top             =   360
            Width           =   15675
            _ExtentX        =   27649
            _ExtentY        =   10980
            _Version        =   393216
            Rows            =   30
            Cols            =   17
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_TipCom 
            Height          =   285
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Comp."
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
         Begin Threed.SSPanel pnl_TipPro 
            Height          =   285
            Left            =   720
            TabIndex        =   14
            Top             =   60
            Width           =   3075
            _Version        =   65536
            _ExtentX        =   5424
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Proceso"
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   285
            Left            =   11400
            TabIndex        =   15
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
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
         Begin Threed.SSPanel pnl_Receptor 
            Height          =   285
            Left            =   6180
            TabIndex        =   16
            Top             =   60
            Width           =   5295
            _Version        =   65536
            _ExtentX        =   9340
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Receptor"
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   285
            Left            =   4860
            TabIndex        =   17
            Top             =   60
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Documento"
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
         Begin Threed.SSPanel pnl_FecEmi 
            Height          =   285
            Left            =   3780
            TabIndex        =   18
            Top             =   60
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fec. Emisión"
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
         Begin Threed.SSPanel pnl_Importe 
            Height          =   285
            Left            =   12480
            TabIndex        =   19
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
         Begin Threed.SSPanel pnl_Selecc 
            Height          =   285
            Left            =   13980
            TabIndex        =   20
            Top             =   60
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "  Selección"
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
            Alignment       =   1
            Begin VB.CheckBox chkSeleccionar 
               BackColor       =   &H00004000&
               Caption         =   "Check1"
               Height          =   255
               Left            =   945
               TabIndex        =   21
               Top             =   10
               Width           =   255
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   11
         Top             =   780
         Width           =   15765
         _Version        =   65536
         _ExtentX        =   27808
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
            Left            =   2310
            Picture         =   "GesCtb_frm_931.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   1740
            Picture         =   "GesCtb_frm_931.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Procesar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1170
            Picture         =   "GesCtb_frm_931.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   600
            Picture         =   "GesCtb_frm_931.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_931.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   15150
            Picture         =   "GesCtb_frm_931.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   585
         Left            =   60
         TabIndex        =   22
         Top             =   1470
         Width           =   15765
         _Version        =   65536
         _ExtentX        =   27808
         _ExtentY        =   1032
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
         Begin VB.ComboBox cmb_Situacion 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   150
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "Mostrar :"
            Height          =   285
            Left            =   150
            TabIndex        =   23
            Top             =   210
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "Frm_Ctb_FacEle_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_RegTot        As Integer
Dim l_int_RegPro        As Integer
Dim l_int_RegSPr        As Integer
Dim l_int_RegErr        As Integer
Dim l_int_InsUpd        As Integer
Dim l_bol_FlgErr        As Boolean
Dim l_str_MsjRef        As String
Dim l_str_RutaLg        As String
Dim l_lng_Codigo        As Long
Dim l_fsobj             As Scripting.FileSystemObject
Dim l_txtStr            As TextStream

   
Private Sub chkSeleccionar_Click()
Dim r_Fila As Integer
   
   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 0 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 7) = ""
         Next r_Fila
      End If
   
      If chkSeleccionar.Value = 1 Then
         For r_Fila = 0 To grd_Listad.Rows - 1
             grd_Listad.TextMatrix(r_Fila, 7) = "X"
         Next r_Fila
      End If
   Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub
Private Sub cmd_Buscar_Click()
   Screen.MousePointer = 11

   Call fs_Buscar
   Call fs_Activa(False)
   Screen.MousePointer = 0
      
  If (grd_Listad.Rows = 0) Then
       Call cmd_Limpia_Click
   End If
End Sub


Private Sub cmd_Editar_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   'CODIGO
   grd_Listad.Col = 9
   moddat_g_str_CodIte = CStr(grd_Listad.Text)
   
   'ESTADO O SITUACION
   grd_Listad.Col = 10
   moddat_g_str_DesIte = CStr(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   If moddat_g_str_DesIte = 2 Then
      Frm_Ctb_FacEle_03.Show 1
   Else
       MsgBox "No es posible modificar, ya que el documento electrónico ya ha sido procesado.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub
Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
   
   r_int_NroFil = 5
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   'CREDITOS INDIRECTOS
   r_obj_Excel.Sheets(1).Name = "CREDITOS INDIRECTOS"
   
   With r_obj_Excel.Sheets(1)
      .Cells(1, 2) = "REPORTE DE DOCUMENTOS ELECTRÓNICOS"
      .Range(.Cells(1, 2), .Cells(1, 16)).Merge
      .Range(.Cells(1, 2), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(1, 16)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 16)).Font.Size = 14
'
'      .Cells(3, 2) = "Reporte al "
'      .Cells(3, 3) = "'" & Format(CDate(ipp_FecIni.Text), "dd/mm/yyyy")
'      .Cells(3, 4) = " AL "
'      .Cells(3, 5) = "'" & Format(CDate(ipp_FecFin.Text), "dd/mm/yyyy")
'      .Cells(3, 4).Font.Bold = True
'      .Range(.Cells(3, 3), .Cells(3, 5)).HorizontalAlignment = xlHAlignCenter
'      .Range(.Cells(3, 2), .Cells(6, 2)).Font.Bold = True
        
      .Cells(r_int_NroFil, 2) = "ITEM"
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
      .Cells(r_int_NroFil, 3) = "TIPO COMPROBANTE"
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
      .Cells(r_int_NroFil, 4) = "TIPO PROCESO" ' CF
      .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
      .Cells(r_int_NroFil, 5) = "FECHA EMISIÓN"
      .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
      .Cells(r_int_NroFil, 6) = "DOCUMENTO"
      .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
      .Cells(r_int_NroFil, 7) = "RAZÓN SOCIAL"
      .Range(.Cells(r_int_NroFil, 7), .Cells(r_int_NroFil + 1, 7)).Merge
      .Cells(r_int_NroFil, 8) = "DIRECCION"
      .Range(.Cells(r_int_NroFil, 8), .Cells(r_int_NroFil + 1, 8)).Merge
      .Cells(r_int_NroFil, 9) = "DISTRITO"
      .Range(.Cells(r_int_NroFil, 9), .Cells(r_int_NroFil + 1, 9)).Merge
      .Cells(r_int_NroFil, 10) = "PROVINCIA"
      .Range(.Cells(r_int_NroFil, 10), .Cells(r_int_NroFil + 1, 10)).Merge
      .Cells(r_int_NroFil, 11) = "DEPARTAMENTO"
      .Range(.Cells(r_int_NroFil, 11), .Cells(r_int_NroFil + 1, 11)).Merge
      .Cells(r_int_NroFil, 12) = "CORREO"
      .Range(.Cells(r_int_NroFil, 12), .Cells(r_int_NroFil + 1, 12)).Merge
      .Cells(r_int_NroFil, 13) = "MONEDA"
      .Range(.Cells(r_int_NroFil, 13), .Cells(r_int_NroFil + 1, 13)).Merge
      .Cells(r_int_NroFil, 14) = "TIPO CAMBIO"
      .Range(.Cells(r_int_NroFil, 14), .Cells(r_int_NroFil + 1, 14)).Merge
      .Cells(r_int_NroFil, 15) = "IMPORTE"
      .Range(.Cells(r_int_NroFil, 15), .Cells(r_int_NroFil + 1, 15)).Merge
      .Cells(r_int_NroFil, 16) = "ESTADO"
      .Range(.Cells(r_int_NroFil, 16), .Cells(r_int_NroFil + 1, 16)).Merge
      
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 16)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 16)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 16)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 16)).VerticalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 10
      .Columns("C").ColumnWidth = 15
      .Columns("D").ColumnWidth = 35
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 13
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 35
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 56
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 25
      .Columns("I").HorizontalAlignment = xlHAlignLeft
      .Columns("J").ColumnWidth = 25
      .Columns("J").HorizontalAlignment = xlHAlignLeft
      .Columns("K").ColumnWidth = 25
      .Columns("K").HorizontalAlignment = xlHAlignLeft
      .Columns("L").ColumnWidth = 40
      .Columns("L").HorizontalAlignment = xlHAlignLeft
      .Columns("M").ColumnWidth = 13.5
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 13.5
      .Columns("N").NumberFormat = "###,###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("O").ColumnWidth = 13.5
      .Columns("O").NumberFormat = "###,###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      .Columns("P").ColumnWidth = 16
      .Columns("P").HorizontalAlignment = xlHAlignCenter
        
      With .Range(.Cells(5, 2), .Cells(6, 16))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
       
      r_int_NroFil = r_int_NroFil + 2
      
      For r_int_NoFlLi = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 2) = r_int_NoFlLi + 1
         .Cells(r_int_NroFil, 3) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 0)
         .Cells(r_int_NroFil, 4) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 1)
         .Cells(r_int_NroFil, 5) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 2)
         .Cells(r_int_NroFil, 6) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 3)
         .Cells(r_int_NroFil, 7) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 4)
         .Cells(r_int_NroFil, 8) = grd_Listad.TextMatrix(r_int_NoFlLi, 11)
         .Cells(r_int_NroFil, 9) = grd_Listad.TextMatrix(r_int_NoFlLi, 12)
         .Cells(r_int_NroFil, 10) = grd_Listad.TextMatrix(r_int_NoFlLi, 13)
         .Cells(r_int_NroFil, 11) = grd_Listad.TextMatrix(r_int_NoFlLi, 14)
         .Cells(r_int_NroFil, 12) = grd_Listad.TextMatrix(r_int_NoFlLi, 15)
         .Cells(r_int_NroFil, 13) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 5)
         .Cells(r_int_NroFil, 14) = grd_Listad.TextMatrix(r_int_NoFlLi, 16)
         .Cells(r_int_NroFil, 15) = grd_Listad.TextMatrix(r_int_NoFlLi, 6)
         If grd_Listad.TextMatrix(r_int_NoFlLi, 10) = 2 Then
            .Cells(r_int_NroFil, 16) = "NO PROCESADO"
         Else
            .Cells(r_int_NroFil, 16) = "PROCESADO"
         End If
         
          r_int_NroFil = r_int_NroFil + 1
      Next r_int_NoFlLi
      
      With .Range(.Cells(7, 2), .Cells(r_int_NroFil, 3))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      With .Range(.Cells(1, 2), .Cells(1, 16))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
      End With
      With .Range(.Cells(5, 2), .Cells(r_int_NroFil - 1, 16))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      With .Range(.Cells(5, 9), .Cells(r_int_NroFil - 1, 9))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      With .Range(.Cells(5, 2), .Cells(6, 16))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).Weight = xlThin
      End With
   End With
   
   r_obj_Excel.Visible = True
End Sub
Private Sub cmd_Limpia_Click()
   grd_Listad.Rows = 0
   chkSeleccionar.Value = 0
'   cmb_Situacion.ListIndex = -1
   Call fs_Activa(True)
End Sub
Private Sub fs_Activa(ByVal p_Activa As Integer)
    cmd_Buscar.Enabled = p_Activa
    cmd_ExpExc.Enabled = Not p_Activa
    cmd_Editar.Enabled = Not p_Activa
    grd_Listad.Enabled = Not p_Activa
    
   If cmb_Situacion.ListIndex = 0 Then
      cmd_Proces.Enabled = Not p_Activa
   ElseIf cmb_Situacion.ListIndex = 1 Then
      cmd_Proces.Enabled = Not p_Activa
   Else
      cmd_Proces.Enabled = Not p_Activa
   End If
   
   cmb_Situacion.Enabled = p_Activa
End Sub
Private Sub cmd_Proces_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer

   'valida seleccion
   r_int_ConSel = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
         r_int_ConSel = r_int_ConSel + 1
      End If
   Next r_int_Contad
   
   If r_int_ConSel = 0 Then
      MsgBox "No se han seleccionados registros para generar documentos electrónicos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'confirma
   If MsgBox("¿Está seguro de generar los documentos electrónicos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Generar_DocEle
   Call fs_Buscar
   Screen.MousePointer = 0
   
   If (grd_Listad.Rows = 0) Then
      Call cmd_Limpia_Click
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub
Private Sub fs_Generar_DocEle()
Dim r_int_Codigo        As Integer
Dim r_int_TipDoc        As Integer
Dim r_str_NumDoc        As String
Dim r_lng_Contad        As Long

   Screen.MousePointer = 11
   moddat_g_int_FlgGrb = 1
'   l_int_RegTot = 0
'   l_int_RegPro = 0
'   l_int_RegSPr = 0
'   l_int_RegErr = 0
'   l_bol_FlgErr = False
    
  
   grd_Listad.Redraw = False
   
'   ReDim l_arr_LogPro(0)
'   ReDim l_arr_LogPro(1)
'   ReDim r_arr_Matriz(0)
   
'   l_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
'   l_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
'   l_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
'   l_arr_LogPro(1).LogPro_NumErr = 0
        
   l_lng_Codigo = 0
   
   For r_lng_Contad = 0 To grd_Listad.Rows - 1
      
      If grd_Listad.TextMatrix(r_lng_Contad, 7) = "X" Then
      
         l_lng_Codigo = Trim(grd_Listad.TextMatrix(r_lng_Contad, 9))
         r_int_TipDoc = Mid(Trim(grd_Listad.TextMatrix(r_lng_Contad, 3)), 1, 1)
         r_str_NumDoc = Mid(Trim(grd_Listad.TextMatrix(r_lng_Contad, 3)), 3)
         
         'Insertar Datos en CNTBL_DOCELE
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         'Inserta en CNTBL_DOCELE
         Do While moddat_g_int_FlgGOK = False
            Call moddat_gs_FecSis
            
            g_str_Parame = "USP_CNTBL_DOCELE ("
            g_str_Parame = g_str_Parame & "" & l_lng_Codigo & ", "                                                'Código
            g_str_Parame = g_str_Parame & CStr(r_int_TipDoc) & ", "                                               'Tipo Documento Receptor
            g_str_Parame = g_str_Parame & "'" & r_str_NumDoc & "', "                                              'Número Documento Receptor
            g_str_Parame = g_str_Parame & 1 & ", "                                                                'insertar a CNTBL_DOCELE
            
            'Datos de Auditoría
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                                       'Código Usuario
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                                        'Nombre Ejecutable
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                                       'Nombre Terminal
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                                       'Código Sucursal
               
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
         
      End If
   Next r_lng_Contad
   
'   l_int_RegPro = r_int_ConAux

   grd_Listad.Redraw = True
      
   Call gs_RefrescaGrid(grd_Listad)
   Call fs_Limpiar
   Call gs_SetFocus(cmd_Buscar)
   
'   fs_Generar_DocEle = True
'   l_int_RegErr = IIf(l_bol_FlgErr = True, 1, 0)
'   l_int_RegSPr = l_int_RegTot - l_int_RegPro - l_int_RegErr
'   l_str_MsjRef = IIf(l_bol_FlgErr = True, "NumRef: " & l_str_MsjRef, "")
End Sub
Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
    
   Call fs_Inicia
   Call fs_Limpiar
   Call fs_Activa(False)
   Call fs_Buscar
     
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub
Private Sub fs_Inicia()

   cmb_Situacion.Clear
   cmb_Situacion.AddItem "NO PROCESADOS"
   cmb_Situacion.AddItem "PROCESADOS"
   cmb_Situacion.AddItem "<< TODOS >>"
   cmb_Situacion.ListIndex = 0
   
   'Datos del Credito
   grd_Listad.ColWidth(0) = 735     'Tipo Comprobante
   grd_Listad.ColWidth(1) = 3000    'Tipo Proceso
   grd_Listad.ColWidth(2) = 1080    'Fecha Emisión
   grd_Listad.ColWidth(3) = 1300    'Documento
   grd_Listad.ColWidth(4) = 5240    'Receptor
   grd_Listad.ColWidth(5) = 1125    'Moneda
   grd_Listad.ColWidth(6) = 1450    'Importe
   grd_Listad.ColWidth(7) = 1300    'Seleccionar
   grd_Listad.ColWidth(8) = 0       'Fecha Emisión sin Formato
   grd_Listad.ColWidth(9) = 0       'Código
   grd_Listad.ColWidth(10) = 0      'Situac
   grd_Listad.ColWidth(11) = 0      'Direccion
   grd_Listad.ColWidth(12) = 0      'Departamento
   grd_Listad.ColWidth(13) = 0      'Provincia
   grd_Listad.ColWidth(14) = 0      'Distrito
   grd_Listad.ColWidth(15) = 0      'Correo
   grd_Listad.ColWidth(16) = 0      'TipCam
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   
   grd_Listad.Rows = 0

End Sub
Private Sub fs_Limpiar()
   Call gs_LimpiaGrid(grd_Listad)
End Sub
Public Sub fs_Buscar()
   
   Call gs_LimpiaGrid(grd_Listad)
   
   'Buscando Información de DocEleTmp
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT DOCELETMP_IDE_TIPDOC, DOCELETMP_TIPPRO           , DOCELETMP_IDE_FECEMI, DOCELETMP_REC_TIPDOC, DOCELETMP_REC_NUMDOC,  "
   g_str_Parame = g_str_Parame & "       DOCELETMP_IDE_TIPMON, DOCELETMP_CAB_TOTVTA_OPEINA, DOCELETMP_CODIGO    , DOCELETMP_SITUAC    , DOCELETMP_REC_DENOMI,  "
   g_str_Parame = g_str_Parame & "       DOCELETMP_REC_DIRCOM, DOCELETMP_REC_DEPART       , DOCELETMP_REC_PROVIN, DOCELETMP_REC_DISTRI, DOCELETMP_REC_CORREC, DOCELETMP_TIPCAM   "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_DOCELETMP "
   
   If cmb_Situacion.ListIndex = 0 Then          'NO PROCESADOS
      g_str_Parame = g_str_Parame & " WHERE DOCELETMP_FECAPR IS NULL "
      g_str_Parame = g_str_Parame & "   AND DOCELETMP_SITUAC = 2 "
   ElseIf cmb_Situacion.ListIndex = 1 Then      'PROCESADOS - FACTURADOR
'      g_str_Parame = g_str_Parame & " WHERE DOCELETMP_FECAPR IS NOT NULL "
      g_str_Parame = g_str_Parame & "  WHERE DOCELETMP_SITUAC = 1 "
   End If
   
   g_str_Parame = g_str_Parame & "    ORDER BY DOCELETMP_CODIGO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      'Tipo de Comprobante
      grd_Listad.Col = 0
      If CStr(g_rst_Princi!DOCELETMP_IDE_TIPDOC) = "01" Then
         grd_Listad.Text = "F"
      ElseIf CStr(g_rst_Princi!DOCELETMP_IDE_TIPDOC) = "03" Then
         grd_Listad.Text = "B"
      ElseIf CStr(g_rst_Princi!DOCELETMP_IDE_TIPDOC) = "07" Then
         grd_Listad.Text = "NC"
      End If
      
      'Tipo de Proceso
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_ParDes("539", CStr(g_rst_Princi!DOCELETMP_TIPPRO))
      
      'Fecha de Emisión
      grd_Listad.Col = 2
      grd_Listad.Text = gf_FormatoFecha(Format(CStr(g_rst_Princi!DOCELETMP_IDE_FECEMI), "YYYYMMDD"))
      
      'Documento
      grd_Listad.Col = 3
      grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_REC_TIPDOC) & "-" & CStr(g_rst_Princi!DOCELETMP_REC_NUMDOC)
            
      'Receptor
      grd_Listad.Col = 4
      grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_REC_DENOMI)
      
      'Moneda
      grd_Listad.Col = 5
      grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_IDE_TIPMON)
     
     'Importe
      grd_Listad.Col = 6
      grd_Listad.Text = Format(g_rst_Princi!DOCELETMP_CAB_TOTVTA_OPEINA, "###,###,###,##0.00")
      
      'Seleccionar
      'grd_Listad.Col = 7
      
      'Fecha Emisión (sin formato)
      grd_Listad.Col = 8
      If IsNull(g_rst_Princi!DOCELETMP_IDE_FECEMI) Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = Format(CStr(g_rst_Princi!DOCELETMP_IDE_FECEMI), "YYYYMMDD")
      End If
      
      'Código
      grd_Listad.Col = 9
      grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_CODIGO)
      
      'Estado
      grd_Listad.Col = 10
      grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_SITUAC)
      
      'Dirección
      grd_Listad.Col = 11
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_DIRCOM) Then
         grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_REC_DIRCOM)
      Else
         grd_Listad.Text = ""
      End If
      
      'Departamento
      grd_Listad.Col = 12
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_DISTRI) Then
         grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_REC_DISTRI)
      Else
         grd_Listad.Text = ""
      End If
      
      'Provincia
      grd_Listad.Col = 13
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_PROVIN) Then
         grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_REC_PROVIN)
      Else
         grd_Listad.Text = ""
      End If
      
      'Distrito
      grd_Listad.Col = 14
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_DEPART) Then
         grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_REC_DEPART)
      Else
         grd_Listad.Text = ""
      End If
      
      'Correo
      grd_Listad.Col = 15
      If Not IsNull(g_rst_Princi!DOCELETMP_REC_CORREC) Then
         grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_REC_CORREC)
      Else
         grd_Listad.Text = ""
      End If
      
      grd_Listad.Col = 16
      If Not IsNull(g_rst_Princi!DOCELETMP_TIPCAM) Then
         grd_Listad.Text = CStr(g_rst_Princi!DOCELETMP_TIPCAM)
      Else
         grd_Listad.Text = 0
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
End Sub
Private Sub fs_Escribir_Linea(p_ruta As String, p_texto As String)
   On Error GoTo MyError

   'Escribir en archivo según se ubique
   Set l_fsobj = New FileSystemObject
   Set l_txtStr = l_fsobj.OpenTextFile(p_ruta, ForAppending, False)
   l_txtStr.WriteLine (p_texto)
   l_txtStr.Close
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Escribir_Linea")
End Sub

Private Sub grd_Listad_DblClick()
 If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 7
      
      If grd_Listad.Text = "X" Then
         grd_Listad.Text = ""
      Else
         grd_Listad.Text = "X"
      End If
      
      Call gs_RefrescaGrid(grd_Listad)
   End If
End Sub

Private Sub pnl_DocIde_Click()
   If Len(Trim(pnl_DocIde.Tag)) = 0 Or pnl_DocIde.Tag = "D" Then
      pnl_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_FecEmi_Click()
   If Len(Trim(pnl_FecEmi.Tag)) = 0 Or pnl_FecEmi.Tag = "D" Then
      pnl_FecEmi.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "C")
   Else
      pnl_FecEmi.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub pnl_Importe_Click()
   If Len(Trim(pnl_Importe.Tag)) = 0 Or pnl_Importe.Tag = "D" Then
      pnl_Importe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Importe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "N-")
   End If
End Sub

Private Sub pnl_Receptor_Click()
   If Len(Trim(pnl_Receptor.Tag)) = 0 Or pnl_Receptor.Tag = "D" Then
      pnl_Receptor.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Receptor.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub
