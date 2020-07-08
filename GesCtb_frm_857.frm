VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptCtb_31 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14760
   Icon            =   "GesCtb_frm_857.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   14760
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel10 
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14760
      _Version        =   65536
      _ExtentX        =   26035
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
         Left            =   600
         TabIndex        =   3
         Top             =   180
         Width           =   4785
         _Version        =   65536
         _ExtentX        =   8440
         _ExtentY        =   529
         _StockProps     =   15
         Caption         =   "Reporte de Morosidad de Cartera Atrasada - Detalle"
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
         Picture         =   "GesCtb_frm_857.frx":000C
         Top             =   90
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   645
      Left            =   0
      TabIndex        =   4
      Top             =   690
      Width           =   14760
      _Version        =   65536
      _ExtentX        =   26035
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
      Font3D          =   2
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   14100
         Picture         =   "GesCtb_frm_857.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   615
      End
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   45
         Picture         =   "GesCtb_frm_857.frx":0758
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exportar a Excel"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   6780
      Left            =   0
      TabIndex        =   6
      Top             =   1350
      Width           =   14745
      _Version        =   65536
      _ExtentX        =   26009
      _ExtentY        =   11959
      _StockProps     =   15
      Caption         =   "SSPanel1"
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   525
         Left            =   90
         TabIndex        =   7
         Top             =   90
         Width           =   14595
         _Version        =   65536
         _ExtentX        =   25744
         _ExtentY        =   926
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
         Begin VB.Label lblConcepto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1185
            TabIndex        =   2
            Top             =   120
            Width           =   13305
         End
         Begin VB.Label Label5 
            Caption         =   "Producto :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   150
            Width           =   1230
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6045
         Left            =   90
         TabIndex        =   9
         Top             =   660
         Width           =   14580
         _Version        =   65536
         _ExtentX        =   25717
         _ExtentY        =   10663
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisCla 
            Height          =   6015
            Left            =   30
            TabIndex        =   10
            Top             =   30
            Width           =   14505
            _ExtentX        =   25585
            _ExtentY        =   10610
            _Version        =   393216
            Rows            =   5
            Cols            =   37
            BackColorSel    =   32768
            ForeColorSel    =   14737632
            FocusRect       =   0
            SelectionMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_Contad     As Integer
Dim r_int_Conta      As Integer
Dim r_int_nrofil     As Integer
Dim r_int_NroIni     As Integer
      
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      'Titulo
      .Cells(2, 1) = "REPORTE DE MOROSIDAD DE CARTERA ATRASADA - DETALLE DEL PRODUCTO " & moddat_g_str_NomPrd
      .Range(.Cells(2, 1), .Cells(2, 37)).Merge
      .Range("A2:Y2").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(4, 1), .Cells(4, 37)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 1), .Cells(4, 37)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 37)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 37)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 37)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 37)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(5, 1), .Cells(5, 37)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(5, 1), .Cells(5, 37)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(5, 1), .Cells(5, 37)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(5, 1), .Cells(5, 37)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(5, 1), .Cells(5, 37)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(5, 1), .Cells(5, 37)).Borders(xlInsideVertical).LineStyle = xlContinuous

      'Primera Linea
      r_int_nrofil = 4
      .Cells(r_int_nrofil, 1) = "Nº"
      .Cells(r_int_nrofil, 2) = "ENERO"
      .Cells(r_int_nrofil, 5) = "FEBRERO"
      .Cells(r_int_nrofil, 8) = "MARZO"
      .Cells(r_int_nrofil, 11) = "ABRIL"
      .Cells(r_int_nrofil, 14) = "MAYO"
      .Cells(r_int_nrofil, 17) = "JUNIO"
      .Cells(r_int_nrofil, 20) = "JULIO"
      .Cells(r_int_nrofil, 23) = "AGOSTO"
      .Cells(r_int_nrofil, 26) = "SETIEMBRE"
      .Cells(r_int_nrofil, 29) = "OCTUBRE"
      .Cells(r_int_nrofil, 32) = "NOVIEMBRE"
      .Cells(r_int_nrofil, 35) = "DICIEMBRE"
      .Range(.Cells(r_int_nrofil, 2), .Cells(r_int_nrofil, 4)).Merge
      .Range(.Cells(r_int_nrofil, 5), .Cells(r_int_nrofil, 7)).Merge
      .Range(.Cells(r_int_nrofil, 8), .Cells(r_int_nrofil, 10)).Merge
      .Range(.Cells(r_int_nrofil, 11), .Cells(r_int_nrofil, 13)).Merge
      .Range(.Cells(r_int_nrofil, 14), .Cells(r_int_nrofil, 16)).Merge
      .Range(.Cells(r_int_nrofil, 17), .Cells(r_int_nrofil, 19)).Merge
      .Range(.Cells(r_int_nrofil, 20), .Cells(r_int_nrofil, 22)).Merge
      .Range(.Cells(r_int_nrofil, 23), .Cells(r_int_nrofil, 25)).Merge
      .Range(.Cells(r_int_nrofil, 26), .Cells(r_int_nrofil, 28)).Merge
      .Range(.Cells(r_int_nrofil, 29), .Cells(r_int_nrofil, 31)).Merge
      .Range(.Cells(r_int_nrofil, 32), .Cells(r_int_nrofil, 34)).Merge
      .Range(.Cells(r_int_nrofil, 35), .Cells(r_int_nrofil, 37)).Merge
      
      'Segunda Linea
      r_int_nrofil = r_int_nrofil + 1
      
      .Columns("A").ColumnWidth = 3:
      .Columns("B").ColumnWidth = 30:  .Cells(r_int_nrofil, 2) = "CLIENTE": .Cells(r_int_nrofil, 2).HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 8:   .Cells(r_int_nrofil, 3) = "VENCIDO": .Cells(r_int_nrofil, 3).HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 5:   .Cells(r_int_nrofil, 4) = "MORA":    .Cells(r_int_nrofil, 4).HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 30:  .Cells(r_int_nrofil, 5) = "CLIENTE": .Cells(r_int_nrofil, 5).HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 8:   .Cells(r_int_nrofil, 6) = "VENCIDO": .Cells(r_int_nrofil, 6).HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 5:   .Cells(r_int_nrofil, 7) = "MORA":    .Cells(r_int_nrofil, 7).HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 30:  .Cells(r_int_nrofil, 8) = "CLIENTE": .Cells(r_int_nrofil, 8).HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 8:   .Cells(r_int_nrofil, 9) = "VENCIDO": .Cells(r_int_nrofil, 9).HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 5:   .Cells(r_int_nrofil, 10) = "MORA":   .Cells(r_int_nrofil, 10).HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 30:  .Cells(r_int_nrofil, 11) = "CLIENTE": .Cells(r_int_nrofil, 11).HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 8:   .Cells(r_int_nrofil, 12) = "VENCIDO": .Cells(r_int_nrofil, 12).HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 5:   .Cells(r_int_nrofil, 13) = "MORA":    .Cells(r_int_nrofil, 13).HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 30:  .Cells(r_int_nrofil, 14) = "CLIENTE": .Cells(r_int_nrofil, 14).HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 8:   .Cells(r_int_nrofil, 15) = "VENCIDO": .Cells(r_int_nrofil, 15).HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 5:   .Cells(r_int_nrofil, 16) = "MORA":    .Cells(r_int_nrofil, 16).HorizontalAlignment = xlHAlignCenter
      .Columns("Q").ColumnWidth = 30:  .Cells(r_int_nrofil, 17) = "CLIENTE": .Cells(r_int_nrofil, 17).HorizontalAlignment = xlHAlignCenter
      .Columns("R").ColumnWidth = 8:   .Cells(r_int_nrofil, 18) = "VENCIDO": .Cells(r_int_nrofil, 18).HorizontalAlignment = xlHAlignCenter
      .Columns("S").ColumnWidth = 5:   .Cells(r_int_nrofil, 19) = "MORA":    .Cells(r_int_nrofil, 19).HorizontalAlignment = xlHAlignCenter
      .Columns("T").ColumnWidth = 30:  .Cells(r_int_nrofil, 20) = "CLIENTE": .Cells(r_int_nrofil, 20).HorizontalAlignment = xlHAlignCenter
      .Columns("U").ColumnWidth = 8:   .Cells(r_int_nrofil, 21) = "VENCIDO": .Cells(r_int_nrofil, 21).HorizontalAlignment = xlHAlignCenter
      .Columns("V").ColumnWidth = 5:   .Cells(r_int_nrofil, 22) = "MORA":    .Cells(r_int_nrofil, 22).HorizontalAlignment = xlHAlignCenter
      .Columns("W").ColumnWidth = 30:  .Cells(r_int_nrofil, 23) = "CLIENTE": .Cells(r_int_nrofil, 23).HorizontalAlignment = xlHAlignCenter
      .Columns("X").ColumnWidth = 8:   .Cells(r_int_nrofil, 24) = "VENCIDO": .Cells(r_int_nrofil, 24).HorizontalAlignment = xlHAlignCenter
      .Columns("Y").ColumnWidth = 5:   .Cells(r_int_nrofil, 25) = "MORA":    .Cells(r_int_nrofil, 25).HorizontalAlignment = xlHAlignCenter
      .Columns("Z").ColumnWidth = 30:  .Cells(r_int_nrofil, 26) = "CLIENTE": .Cells(r_int_nrofil, 26).HorizontalAlignment = xlHAlignCenter
      .Columns("AA").ColumnWidth = 8:  .Cells(r_int_nrofil, 27) = "VENCIDO": .Cells(r_int_nrofil, 27).HorizontalAlignment = xlHAlignCenter
      .Columns("AB").ColumnWidth = 5:  .Cells(r_int_nrofil, 28) = "MORA":    .Cells(r_int_nrofil, 28).HorizontalAlignment = xlHAlignCenter
      .Columns("AC").ColumnWidth = 30: .Cells(r_int_nrofil, 29) = "CLIENTE": .Cells(r_int_nrofil, 29).HorizontalAlignment = xlHAlignCenter
      .Columns("AD").ColumnWidth = 8:  .Cells(r_int_nrofil, 30) = "VENCIDO": .Cells(r_int_nrofil, 30).HorizontalAlignment = xlHAlignCenter
      .Columns("AE").ColumnWidth = 5:  .Cells(r_int_nrofil, 31) = "MORA":    .Cells(r_int_nrofil, 31).HorizontalAlignment = xlHAlignCenter
      .Columns("AF").ColumnWidth = 30: .Cells(r_int_nrofil, 32) = "CLIENTE": .Cells(r_int_nrofil, 32).HorizontalAlignment = xlHAlignCenter
      .Columns("AG").ColumnWidth = 8:  .Cells(r_int_nrofil, 33) = "VENCIDO": .Cells(r_int_nrofil, 33).HorizontalAlignment = xlHAlignCenter
      .Columns("AH").ColumnWidth = 5:  .Cells(r_int_nrofil, 34) = "MORA":    .Cells(r_int_nrofil, 34).HorizontalAlignment = xlHAlignCenter
      .Columns("AI").ColumnWidth = 30: .Cells(r_int_nrofil, 35) = "CLIENTE": .Cells(r_int_nrofil, 35).HorizontalAlignment = xlHAlignCenter
      .Columns("AJ").ColumnWidth = 8:  .Cells(r_int_nrofil, 36) = "VENCIDO": .Cells(r_int_nrofil, 36).HorizontalAlignment = xlHAlignCenter
      .Columns("AK").ColumnWidth = 5:  .Cells(r_int_nrofil, 37) = "MORA":    .Cells(r_int_nrofil, 37).HorizontalAlignment = xlHAlignCenter
      
      'Combina celdas de primer linea
      .Range("A4:A4").HorizontalAlignment = xlHAlignCenter
      .Range("B4:D4").HorizontalAlignment = xlHAlignCenter
      .Range("E4:G4").HorizontalAlignment = xlHAlignCenter
      .Range("H4:J4").HorizontalAlignment = xlHAlignCenter
      .Range("K4:M4").HorizontalAlignment = xlHAlignCenter
      .Range("N4:P4").HorizontalAlignment = xlHAlignCenter
      .Range("Q4:S4").HorizontalAlignment = xlHAlignCenter
      .Range("T4:V4").HorizontalAlignment = xlHAlignCenter
      .Range("W4:Y4").HorizontalAlignment = xlHAlignCenter
      .Range("Z4:AB4").HorizontalAlignment = xlHAlignCenter
      .Range("AC4:AE4").HorizontalAlignment = xlHAlignCenter
      .Range("AF4:AH4").HorizontalAlignment = xlHAlignCenter
      .Range("AI4:AK4").HorizontalAlignment = xlHAlignCenter
      
      'Formatea titulo
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 37)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 37)).Font.Size = 11
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 37)).Font.Bold = True
      
      'Exporta filas
      r_int_NroIni = 2
      For r_int_Contad = r_int_NroIni To grd_LisCla.Rows - 1
         .Cells(r_int_Contad + 4, 1) = r_int_NroIni - 1
         .Cells(r_int_Contad + 4, 2) = grd_LisCla.TextMatrix(r_int_NroIni, 1)
         .Cells(r_int_Contad + 4, 3) = grd_LisCla.TextMatrix(r_int_NroIni, 2)
         .Cells(r_int_Contad + 4, 4) = grd_LisCla.TextMatrix(r_int_NroIni, 3)
         .Cells(r_int_Contad + 4, 5) = grd_LisCla.TextMatrix(r_int_NroIni, 4)
         .Cells(r_int_Contad + 4, 6) = grd_LisCla.TextMatrix(r_int_NroIni, 5)
         .Cells(r_int_Contad + 4, 7) = grd_LisCla.TextMatrix(r_int_NroIni, 6)
         .Cells(r_int_Contad + 4, 8) = grd_LisCla.TextMatrix(r_int_NroIni, 7)
         .Cells(r_int_Contad + 4, 9) = grd_LisCla.TextMatrix(r_int_NroIni, 8)
         .Cells(r_int_Contad + 4, 10) = grd_LisCla.TextMatrix(r_int_NroIni, 9)
         .Cells(r_int_Contad + 4, 11) = grd_LisCla.TextMatrix(r_int_NroIni, 10)
         .Cells(r_int_Contad + 4, 12) = grd_LisCla.TextMatrix(r_int_NroIni, 11)
         .Cells(r_int_Contad + 4, 13) = grd_LisCla.TextMatrix(r_int_NroIni, 12)
         .Cells(r_int_Contad + 4, 14) = grd_LisCla.TextMatrix(r_int_NroIni, 13)
         .Cells(r_int_Contad + 4, 15) = grd_LisCla.TextMatrix(r_int_NroIni, 14)
         .Cells(r_int_Contad + 4, 16) = grd_LisCla.TextMatrix(r_int_NroIni, 15)
         .Cells(r_int_Contad + 4, 17) = grd_LisCla.TextMatrix(r_int_NroIni, 16)
         .Cells(r_int_Contad + 4, 18) = grd_LisCla.TextMatrix(r_int_NroIni, 17)
         .Cells(r_int_Contad + 4, 19) = grd_LisCla.TextMatrix(r_int_NroIni, 18)
         .Cells(r_int_Contad + 4, 20) = grd_LisCla.TextMatrix(r_int_NroIni, 19)
         .Cells(r_int_Contad + 4, 21) = grd_LisCla.TextMatrix(r_int_NroIni, 20)
         .Cells(r_int_Contad + 4, 22) = grd_LisCla.TextMatrix(r_int_NroIni, 21)
         .Cells(r_int_Contad + 4, 23) = grd_LisCla.TextMatrix(r_int_NroIni, 22)
         .Cells(r_int_Contad + 4, 24) = grd_LisCla.TextMatrix(r_int_NroIni, 23)
         .Cells(r_int_Contad + 4, 25) = grd_LisCla.TextMatrix(r_int_NroIni, 24)
         .Cells(r_int_Contad + 4, 26) = grd_LisCla.TextMatrix(r_int_NroIni, 25)
         .Cells(r_int_Contad + 4, 27) = grd_LisCla.TextMatrix(r_int_NroIni, 26)
         .Cells(r_int_Contad + 4, 28) = grd_LisCla.TextMatrix(r_int_NroIni, 27)
         .Cells(r_int_Contad + 4, 29) = grd_LisCla.TextMatrix(r_int_NroIni, 28)
         .Cells(r_int_Contad + 4, 30) = grd_LisCla.TextMatrix(r_int_NroIni, 29)
         .Cells(r_int_Contad + 4, 31) = grd_LisCla.TextMatrix(r_int_NroIni, 30)
         .Cells(r_int_Contad + 4, 32) = grd_LisCla.TextMatrix(r_int_NroIni, 31)
         .Cells(r_int_Contad + 4, 33) = grd_LisCla.TextMatrix(r_int_NroIni, 32)
         .Cells(r_int_Contad + 4, 34) = grd_LisCla.TextMatrix(r_int_NroIni, 33)
         .Cells(r_int_Contad + 4, 35) = grd_LisCla.TextMatrix(r_int_NroIni, 34)
         .Cells(r_int_Contad + 4, 36) = grd_LisCla.TextMatrix(r_int_NroIni, 35)
         .Cells(r_int_Contad + 4, 37) = grd_LisCla.TextMatrix(r_int_NroIni, 36)
        
         r_int_NroIni = r_int_NroIni + 1
         .Range(.Cells(1, 1), .Cells(r_int_Contad + 4, 37)).Font.Name = "Arial"
         .Range(.Cells(1, 1), .Cells(r_int_Contad + 4, 37)).Font.Size = 8
         
         For r_int_Conta = 1 To 38
            .Range(.Cells(r_int_Contad + 4, r_int_Conta), .Cells(r_int_Contad + 4, r_int_Conta)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Next
         .Range(.Cells(r_int_Contad + 4, 1), .Cells(r_int_Contad + 4, 37)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      Next
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmd_ExpExc_Click()
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
   
   Call fs_Inicia
   Call fs_Buscar_Morosidad_Cartera
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   lblConcepto.Caption = moddat_g_str_NomPrd
   Call gs_LimpiaGrid(grd_LisCla)
   
   grd_LisCla.AllowUserResizing = flexResizeColumns
   grd_LisCla.SelectionMode = flexSelectionByRow
   grd_LisCla.FocusRect = flexFocusNone
   grd_LisCla.HighLight = flexHighlightAlways
   
   grd_LisCla.ColWidth(0) = 450
   grd_LisCla.ColWidth(1) = 3200
   grd_LisCla.ColWidth(2) = 900
   grd_LisCla.ColWidth(3) = 550
   grd_LisCla.ColWidth(4) = 3200
   grd_LisCla.ColWidth(5) = 900
   grd_LisCla.ColWidth(6) = 550
   grd_LisCla.ColWidth(7) = 3200
   grd_LisCla.ColWidth(8) = 900
   grd_LisCla.ColWidth(9) = 550
   grd_LisCla.ColWidth(10) = 3200
   grd_LisCla.ColWidth(11) = 900
   grd_LisCla.ColWidth(12) = 550
   grd_LisCla.ColWidth(13) = 3200
   grd_LisCla.ColWidth(14) = 900
   grd_LisCla.ColWidth(15) = 550
   grd_LisCla.ColWidth(16) = 3200
   grd_LisCla.ColWidth(17) = 900
   grd_LisCla.ColWidth(18) = 550
   grd_LisCla.ColWidth(19) = 3200
   grd_LisCla.ColWidth(20) = 900
   grd_LisCla.ColWidth(21) = 550
   grd_LisCla.ColWidth(22) = 3200
   grd_LisCla.ColWidth(23) = 900
   grd_LisCla.ColWidth(24) = 550
   grd_LisCla.ColWidth(25) = 3200
   grd_LisCla.ColWidth(26) = 900
   grd_LisCla.ColWidth(27) = 550
   grd_LisCla.ColWidth(28) = 3200
   grd_LisCla.ColWidth(29) = 900
   grd_LisCla.ColWidth(30) = 550
   grd_LisCla.ColWidth(31) = 3200
   grd_LisCla.ColWidth(32) = 900
   grd_LisCla.ColWidth(33) = 550
   grd_LisCla.ColWidth(34) = 3200
   grd_LisCla.ColWidth(35) = 900
   grd_LisCla.ColWidth(36) = 550
   grd_LisCla.ColAlignment(0) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(1) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(2) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(3) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(4) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(5) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(6) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(7) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(8) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(9) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(10) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(11) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(12) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(13) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(14) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(15) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(16) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(17) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(18) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(19) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(20) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(21) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(22) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(23) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(24) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(25) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(26) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(27) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(28) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(29) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(30) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(31) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(32) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(33) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(34) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(35) = flexAlignCenterCenter
   grd_LisCla.ColAlignment(36) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar_Morosidad_Cartera()
Dim r_int_NumCol     As Integer
Dim r_int_NumFil     As Integer
Dim r_int_ContMes    As Integer
Dim r_int_Contad     As Integer
Dim r_int_TotCont    As Integer
Dim r_int_Conta      As Integer
Dim r_str_cadInner   As String
Dim r_str_cadWhere   As String
Dim r_str_CodPrd     As String

   grd_LisCla.Redraw = False
   Call gs_LimpiaGrid(grd_LisCla)
   
   'Fila 0
   grd_LisCla.Rows = 0
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   
   grd_LisCla.Col = 0:   grd_LisCla.Text = "Nº"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "ENERO"
   grd_LisCla.Col = 2:   grd_LisCla.Text = "ENERO"
   grd_LisCla.Col = 3:   grd_LisCla.Text = "ENERO"
   grd_LisCla.Col = 4:   grd_LisCla.Text = "FEBRERO"
   grd_LisCla.Col = 5:   grd_LisCla.Text = "FEBRERO"
   grd_LisCla.Col = 6:   grd_LisCla.Text = "FEBRERO"
   grd_LisCla.Col = 7:   grd_LisCla.Text = "MARZO"
   grd_LisCla.Col = 8:   grd_LisCla.Text = "MARZO"
   grd_LisCla.Col = 9:   grd_LisCla.Text = "MARZO"
   grd_LisCla.Col = 10:  grd_LisCla.Text = "ABRIL"
   grd_LisCla.Col = 11:  grd_LisCla.Text = "ABRIL"
   grd_LisCla.Col = 12:  grd_LisCla.Text = "ABRIL"
   grd_LisCla.Col = 13:  grd_LisCla.Text = "MAYO"
   grd_LisCla.Col = 14:  grd_LisCla.Text = "MAYO"
   grd_LisCla.Col = 15:  grd_LisCla.Text = "MAYO"
   grd_LisCla.Col = 16:  grd_LisCla.Text = "JUNIO"
   grd_LisCla.Col = 17:  grd_LisCla.Text = "JUNIO"
   grd_LisCla.Col = 18:  grd_LisCla.Text = "JUNIO"
   grd_LisCla.Col = 19:  grd_LisCla.Text = "JULIO"
   grd_LisCla.Col = 20:  grd_LisCla.Text = "JULIO"
   grd_LisCla.Col = 21:  grd_LisCla.Text = "JULIO"
   grd_LisCla.Col = 22:  grd_LisCla.Text = "AGOSTO"
   grd_LisCla.Col = 23:  grd_LisCla.Text = "AGOSTO"
   grd_LisCla.Col = 24:  grd_LisCla.Text = "AGOSTO"
   grd_LisCla.Col = 25:  grd_LisCla.Text = "SETIEMBRE"
   grd_LisCla.Col = 26:  grd_LisCla.Text = "SETIEMBRE"
   grd_LisCla.Col = 27:  grd_LisCla.Text = "SETIEMBRE"
   grd_LisCla.Col = 28:  grd_LisCla.Text = "OCTUBRE"
   grd_LisCla.Col = 29:  grd_LisCla.Text = "OCTUBRE"
   grd_LisCla.Col = 30:  grd_LisCla.Text = "OCTUBRE"
   grd_LisCla.Col = 31:  grd_LisCla.Text = "NOVIEMBRE"
   grd_LisCla.Col = 32:  grd_LisCla.Text = "NOVIEMBRE"
   grd_LisCla.Col = 33:  grd_LisCla.Text = "NOVIEMBRE"
   grd_LisCla.Col = 34:  grd_LisCla.Text = "DICIEMBRE"
   grd_LisCla.Col = 35:  grd_LisCla.Text = "DICIEMBRE"
   grd_LisCla.Col = 36:  grd_LisCla.Text = "DICIEMBRE"
   
   'Fila 1
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = ""
   grd_LisCla.Col = 1:   grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 2:   grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 3:   grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 4:   grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 5:   grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 6:   grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 7:   grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 8:   grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 9:   grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 10:  grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 11:  grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 12:  grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 13:  grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 14:  grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 15:  grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 16:  grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 17:  grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 18:  grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 19:  grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 20:  grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 21:  grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 22:  grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 23:  grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 24:  grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 25:  grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 26:  grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 27:  grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 28:  grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 29:  grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 30:  grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 31:  grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 32:  grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 33:  grd_LisCla.Text = "MORA"
   grd_LisCla.Col = 34:  grd_LisCla.Text = "CLIENTE"
   grd_LisCla.Col = 35:  grd_LisCla.Text = "VENCIDO"
   grd_LisCla.Col = 36:  grd_LisCla.Text = "MORA"
     
   With grd_LisCla
      .Rows = .Rows + 1
      .MergeCells = flexMergeFree
      .MergeCol(1) = True
      .MergeRow(0) = True
      .FixedCols = 1
      .FixedRows = 2
   End With
         
   r_str_cadInner = ""
   r_str_cadWhere = ""
   If (moddat_g_int_TipRep = 5) Then 'MICROEMPRESARIO
       r_str_cadWhere = "   AND UPPER(TRIM(X.SUBPRD_DESCRI)) LIKE UPPER(TRIM('%MICROEMPRESARIO%')) "
   End If
   
   For r_int_ContMes = 1 To 12
   
      If moddat_g_str_CodMes = r_int_ContMes Then
         If (moddat_g_int_TipRep = 5) Then
             r_str_cadInner = " INNER JOIN CRE_SUBPRD X ON X.SUBPRD_CODPRD = HIPMAE_CODPRD AND X.SUBPRD_CODSUB = HIPMAE_CODSUB AND X.SUBPRD_CODSUB = HIPMAE_CODSUB "
         End If
         
         r_str_CodPrd = ""
         If (moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 2 Or moddat_g_int_TipRep = 5) Then
             Select Case moddat_g_str_TipPar
                    Case 1: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('001')"
                    Case 2: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('002','006','011')"
                    Case 3: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('003')"
                    Case 4: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025')" ','019','021','022','023','024'
                    Case 5: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('019')"
                    Case 6: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('021','022','023')"
                    Case 7: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('024')"
             End Select
         ElseIf (moddat_g_int_TipRep = 3) Then 'MIVIVIENDA
             Select Case moddat_g_str_TipPar
                    Case 1: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025') " ''019','021','022','023','024',
                    Case 3: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('003') "
                    Case 5: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('019') "
                    Case 6: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('021','022','023') "
                    Case 7: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('024') "
                    Case Else: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('003','004','007','009','010','013','014','015','016','017','018','019','021','022','023','024','025') "
             End Select
         ElseIf (moddat_g_int_TipRep = 4) Then 'MICASITA
             Select Case moddat_g_str_TipPar
                    Case 1: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('002','006','011') "
                    Case 4: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('001') "
                    Case Else: r_str_CodPrd = " AND HIPMAE_CODPRD IN ('001','002','006','011') "
             End Select
         End If
   
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ', ' || TRIM(DATGEN_NOMBRE) CLIENTE,"
         g_str_Parame = g_str_Parame & "       HIPMAE_DIAMOR HIPCIE_DIAMOR, "
         g_str_Parame = g_str_Parame & "       NVL((SELECT SUM(HIPCUO_CAPITA+HIPCUO_CAPBBP) "
         g_str_Parame = g_str_Parame & "              FROM CRE_HIPCUO "
         g_str_Parame = g_str_Parame & "             WHERE HIPCUO_NUMOPE = HIPMAE_NUMOPE AND HIPCUO_TIPCRO = 1 AND HIPCUO_SITUAC = 2"
         g_str_Parame = g_str_Parame & "               AND HIPCUO_FECVCT <= '" & moddat_g_str_FecCan & "'"
         g_str_Parame = g_str_Parame & "               AND (TRUNC(TO_DATE('" & moddat_g_str_FecFin & "','YYYYMMDD')) - TRUNC(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD')) > " & moddat_g_int_OrdAct & ") "
         g_str_Parame = g_str_Parame & "               AND (TRUNC(TO_DATE('" & moddat_g_str_FecFin & "','YYYYMMDD')) - TRUNC(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD')) <=90)) ,0) AS HIPCIE_CAPVEN "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
         g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN B ON A.HIPMAE_TDOCLI = B.DATGEN_TIPDOC AND A.HIPMAE_NDOCLI = B.DATGEN_NUMDOC "
         g_str_Parame = g_str_Parame & r_str_cadInner
         g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 2 "
         g_str_Parame = g_str_Parame & r_str_cadWhere
         
         g_str_Parame = g_str_Parame & r_str_CodPrd
         g_str_Parame = g_str_Parame & "   AND HIPMAE_NUMOPE NOT IN (SELECT HIPMAE_NUMOPE AS NRO_OPERACION FROM CRE_HIPMAE"
         g_str_Parame = g_str_Parame & "                              WHERE HIPMAE_SITUAC = 2 "
         
         g_str_Parame = g_str_Parame & r_str_CodPrd
         g_str_Parame = g_str_Parame & "                                AND (SELECT SUM(HIPCUO_CAPITA+HIPCUO_CAPBBP) FROM CRE_HIPCUO"
         g_str_Parame = g_str_Parame & "                                      WHERE HIPCUO_NUMOPE = HIPMAE_NUMOPE AND HIPCUO_TIPCRO = 1"
         g_str_Parame = g_str_Parame & "                                        AND HIPCUO_SITUAC = 2 AND HIPCUO_FECVCT <= '" & moddat_g_str_FecCan & "'"
         g_str_Parame = g_str_Parame & "                                        AND (TRUNC(TO_DATE('" & moddat_g_str_FecFin & "','YYYYMMDD')) - TRUNC(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD')) >90)) > 0)"
         g_str_Parame = g_str_Parame & "   AND (SELECT SUM(HIPCUO_CAPITA+HIPCUO_CAPBBP) "
         g_str_Parame = g_str_Parame & "          FROM CRE_HIPCUO"
         g_str_Parame = g_str_Parame & "         WHERE HIPCUO_NUMOPE = HIPMAE_NUMOPE AND HIPCUO_TIPCRO = 1 AND HIPCUO_SITUAC = 2"
         g_str_Parame = g_str_Parame & "           AND HIPCUO_FECVCT <= '" & moddat_g_str_FecCan & "'"
         g_str_Parame = g_str_Parame & "           AND (TRUNC(TO_DATE('" & moddat_g_str_FecFin & "','YYYYMMDD')) - TRUNC(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD')) > " & moddat_g_int_OrdAct & ")"
         g_str_Parame = g_str_Parame & "           AND (TRUNC(TO_DATE('" & moddat_g_str_FecFin & "','YYYYMMDD')) - TRUNC(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD')) <=90)) > 0 "
         g_str_Parame = g_str_Parame & " UNION "
         g_str_Parame = g_str_Parame & "SELECT TRIM(DATGEN_APEPAT) || ' ' || TRIM(DATGEN_APEMAT) || ', ' || TRIM(DATGEN_NOMBRE) CLIENTE,"
         g_str_Parame = g_str_Parame & "       HIPMAE_DIAMOR HIPCIE_DIAMOR,"
         g_str_Parame = g_str_Parame & "       (SELECT MAX(HIPCUO_CAPITA+HIPCUO_SALCAP)+HIPMAE_SALCON FROM CRE_HIPCUO "
         g_str_Parame = g_str_Parame & "         WHERE HIPCUO_NUMOPE = HIPMAE_NUMOPE AND HIPCUO_TIPCRO = 1 AND HIPCUO_SITUAC = 2)"
         g_str_Parame = g_str_Parame & " +     (SELECT SUM(HIPCUO_CAPBBP) FROM CRE_HIPCUO"
         g_str_Parame = g_str_Parame & "         WHERE HIPCUO_NUMOPE = HIPMAE_NUMOPE AND HIPCUO_TIPCRO = 1 AND HIPCUO_SITUAC = 2) AS HIPCIE_CAPVEN "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A"
         g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN B ON A.HIPMAE_TDOCLI = B.DATGEN_TIPDOC AND A.HIPMAE_NDOCLI = B.DATGEN_NUMDOC"
         g_str_Parame = g_str_Parame & r_str_cadInner
         g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 2"
         g_str_Parame = g_str_Parame & r_str_cadWhere
         
         g_str_Parame = g_str_Parame & r_str_CodPrd
         g_str_Parame = g_str_Parame & "   AND (SELECT SUM(HIPCUO_CAPITA+HIPCUO_CAPBBP) FROM CRE_HIPCUO"
         g_str_Parame = g_str_Parame & "         WHERE HIPCUO_NUMOPE = HIPMAE_NUMOPE AND HIPCUO_TIPCRO = 1 AND HIPCUO_SITUAC = 2 AND HIPCUO_FECVCT <= '" & moddat_g_str_FecCan & "'"
         g_str_Parame = g_str_Parame & "           AND (TRUNC(TO_DATE('" & moddat_g_str_FecFin & "','YYYYMMDD')) - TRUNC(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD')) >90)) > 0"
         g_str_Parame = g_str_Parame & " ORDER BY HIPCIE_DIAMOR DESC "
      Else
         If (moddat_g_int_TipRep = 5) Then
             r_str_cadInner = " INNER JOIN CRE_SUBPRD X ON X.SUBPRD_CODPRD = HIPCIE_CODPRD AND X.SUBPRD_CODSUB = HIPCIE_CODSUB AND X.SUBPRD_CODSUB = HIPCIE_CODSUB "
         End If
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS CLIENTE, "
         g_str_Parame = g_str_Parame & "       HIPCIE_CAPVEN, HIPCIE_DIAMOR "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE A "
         g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.HIPCIE_TDOCLI AND B.DATGEN_NUMDOC = A.HIPCIE_NDOCLI "
         g_str_Parame = g_str_Parame & r_str_cadInner
         g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & r_int_ContMes & ""
         g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & moddat_g_str_CodAno & ""
         g_str_Parame = g_str_Parame & r_str_cadWhere
         If (moddat_g_int_TipRep = 1 Or moddat_g_int_TipRep = 2 Or moddat_g_int_TipRep = 5) Then
             Select Case moddat_g_str_TipPar
                    Case 1: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('001')"
                    Case 2: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('002','006','011')"
                    Case 3: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('003')"
                    Case 4: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025')" ','019','021','022','023','024'
                    Case 5: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('019')"
                    Case 6: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('021','022','023')"
                    Case 7: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('024')"
             End Select
         ElseIf (moddat_g_int_TipRep = 3) Then 'MIVIVIENDA
             Select Case moddat_g_str_TipPar
                    Case 1: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('004','007','009','010','012','013','014','015','016','017','018','025') " ''019','021','022','023','024',
                    Case 3: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('003') "
                    Case 5: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('019') "
                    Case 6: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('021','022','023') "
                    Case 7: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('024') "
                    Case Else: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('003','004','007','009','010','012','013','014','015','016','017','018','019','021','022','023','024','025') "
             End Select
         ElseIf (moddat_g_int_TipRep = 4) Then 'MICASITA
             Select Case moddat_g_str_TipPar
                    Case 1: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('002','006','011') "
                    Case 4: g_str_Parame = g_str_Parame & " AND HIPCIE_CODPRD IN ('001') "
                    Case Else: g_str_Parame = g_str_Parame & "   AND HIPCIE_CODPRD IN ('001','002','006','011') "
             End Select
         End If
         g_str_Parame = g_str_Parame & "   AND HIPCIE_DIAMOR > " & moddat_g_int_OrdAct
         g_str_Parame = g_str_Parame & " ORDER BY HIPCIE_DIAMOR DESC"
      End If
      
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "Error al ejecutar la consulta de Cartera de Morosidad.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      r_int_NumCol = (r_int_ContMes * 3) - 3
      r_int_NumFil = 2
      r_int_Contad = 0
      r_int_Conta = 1
      
      'Carga grid
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            grd_LisCla.Row = r_int_NumFil
            
            grd_LisCla.Col = 0
            grd_LisCla.CellAlignment = flexAlignRightCenter
            grd_LisCla.Text = r_int_Conta
         
            grd_LisCla.Col = r_int_NumCol + 1
            grd_LisCla.CellAlignment = flexAlignLeftCenter
            grd_LisCla.Text = g_rst_Princi!CLIENTE
            
            grd_LisCla.Col = r_int_NumCol + 2
            grd_LisCla.CellAlignment = flexAlignRightCenter
            grd_LisCla.Text = Format(g_rst_Princi!HIPCIE_CAPVEN, "###,##.00")
            
            grd_LisCla.Col = r_int_NumCol + 3
            grd_LisCla.CellAlignment = flexAlignCenterCenter
            grd_LisCla.Text = g_rst_Princi!HIPCIE_DIAMOR
            
            r_int_NumFil = r_int_NumFil + 1
            r_int_Contad = r_int_Contad + 1
            r_int_Conta = r_int_Conta + 1
            grd_LisCla.Rows = grd_LisCla.Rows + 1
           
            g_rst_Princi.MoveNext
         Loop

         If r_int_Contad >= r_int_TotCont Then
            r_int_TotCont = r_int_Contad
         End If
         
         r_int_NumCol = r_int_NumCol + 1
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next
   
   grd_LisCla.Rows = r_int_TotCont + 2
   grd_LisCla.Redraw = True
     
   If (grd_LisCla.Rows > 2) Then
       Call gs_UbicaGrid(grd_LisCla, 2)
   End If
End Sub

Private Sub grd_LisCla_SelChange()
   If grd_LisCla.Rows > 2 Then
      grd_LisCla.RowSel = grd_LisCla.Row
   End If
End Sub
