VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_RegCom_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   Icon            =   "GesCtb_frm_186.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8205
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9617
      _Version        =   65536
      _ExtentX        =   16960
      _ExtentY        =   14473
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   5835
         Left            =   60
         TabIndex        =   9
         Top             =   1470
         Width           =   9465
         _Version        =   65536
         _ExtentX        =   16695
         _ExtentY        =   10292
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
            Height          =   5415
            Left            =   30
            TabIndex        =   10
            Top             =   360
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   9551
            _Version        =   393216
            Rows            =   21
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_NumDoc 
            Height          =   285
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   1995
            _Version        =   65536
            _ExtentX        =   3519
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Documento"
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
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   285
            Left            =   2040
            TabIndex        =   15
            Top             =   60
            Width           =   7035
            _Version        =   65536
            _ExtentX        =   12409
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Razón Social"
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   9465
         _Version        =   65536
         _ExtentX        =   16695
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
            Height          =   495
            Left            =   570
            TabIndex        =   13
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Proveedores"
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
            Left            =   30
            Picture         =   "GesCtb_frm_186.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   14
         Top             =   780
         Width           =   9465
         _Version        =   65536
         _ExtentX        =   16695
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
            Left            =   2460
            Picture         =   "GesCtb_frm_186.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   1860
            Picture         =   "GesCtb_frm_186.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   8850
            Picture         =   "GesCtb_frm_186.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1260
            Picture         =   "GesCtb_frm_186.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Eliminar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_186.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_186.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Adicionar"
            Top             =   30
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   630
         Left            =   60
         TabIndex        =   16
         Top             =   7350
         Width           =   9465
         _Version        =   65536
         _ExtentX        =   16695
         _ExtentY        =   1111
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
         Begin VB.ComboBox cmb_Buscar 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   2595
         End
         Begin VB.TextBox txt_Buscar 
            Height          =   315
            Left            =   5160
            MaxLength       =   100
            TabIndex        =   1
            Top             =   180
            Width           =   4155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Por:"
            Height          =   195
            Left            =   4290
            TabIndex        =   18
            Top             =   240
            Width           =   825
         End
         Begin VB.Label lbl_NomEti 
            AutoSize        =   -1  'True
            Caption         =   "Columna a Buscar:"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_RegCom_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_arr_Sucurs()      As moddat_tpo_Genera
Dim r_str_Origen        As String
Dim l_var_ColAnt        As Variant

Private Sub cmb_Buscar_Click()
   If (cmb_Buscar.ListIndex = 0 Or cmb_Buscar.ListIndex = -1) Then
       txt_Buscar.Enabled = False
        Call fs_BuscarProv
   Else
       txt_Buscar.Enabled = True
   End If
   txt_Buscar.Text = ""
End Sub

Private Sub cmb_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_Buscar.Enabled = False) Then
          Call gs_SetFocus(cmd_Agrega)
      Else
          Call gs_SetFocus(txt_Buscar)
      End If
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   frm_Ctb_RegCom_02.Show 1
End Sub

Private Sub cmd_Borrar_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Seguro que desea eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_TipDoc = CInt(grd_Listad.Text)
   grd_Listad.Col = 1
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   Call gs_RefrescaGrid(grd_Listad)
   
   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CNTBL_MAEPRV_BORRAR ( "
   g_str_Parame = g_str_Parame & Trim(moddat_g_str_TipDoc) & ", " 'MAEPRV_TIPDOC
   g_str_Parame = g_str_Parame & "'" & Trim(moddat_g_str_NumDoc) & "', " 'MAEPRV_NUMDOC
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo completar la eliminación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Screen.MousePointer = 0
      Exit Sub
   Else
      MsgBox "El proveedor se elimino correctamente.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
   
   Call fs_BuscarProv
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Consul_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   Call gs_RefrescaGrid(grd_Listad)
   
   grd_Listad.Col = 0
   moddat_g_str_TipDoc = CInt(grd_Listad.Text)
   grd_Listad.Col = 1
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 0
   frm_Ctb_RegCom_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_str_TipDoc = ""
   moddat_g_str_NumDoc = ""
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_TipDoc = CInt(grd_Listad.Text)
   grd_Listad.Col = 1
   moddat_g_str_NumDoc = CStr(grd_Listad.Text)
      
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb = 2
   frm_Ctb_RegCom_02.Show 1
   
   'Call fs_BuscarProv
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NumFil        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE PROVEEDORES"
      .Range(.Cells(2, 2), .Cells(2, 8)).Merge
      .Range(.Cells(2, 2), .Cells(2, 8)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 8)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(3, 2) = "NRO DOCUMENTO"
      .Cells(3, 3) = "RAZON SOCIAL"
      .Cells(3, 4) = "TIPO CONTRIBUYENTE"
      .Cells(3, 5) = "CONDICION"
      .Cells(3, 6) = "CUENTA DETRACCION(BCO NAC.)"
      .Cells(3, 7) = "TIPO PERSONAL"
      .Cells(3, 8) = "CODIGO PLANILLA"
      
      .Range(.Cells(3, 2), .Cells(3, 8)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 8)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 18
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 44
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 20
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 11
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 30
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 17
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      
      .Range(.Cells(1, 1), .Cells(10, 8)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 8)).Font.Size = 11
      
      r_int_NumFil = 2
      For r_int_Contad = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NumFil + 2, 2) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 2, 2) 'ID-PROVEEDOR
         .Cells(r_int_NumFil + 2, 3) = grd_Listad.TextMatrix(r_int_NumFil - 2, 3) 'RAZONSOCIAL
         .Cells(r_int_NumFil + 2, 4) = grd_Listad.TextMatrix(r_int_NumFil - 2, 4) 'TIPO CONTRIBUYENTE
         .Cells(r_int_NumFil + 2, 5) = grd_Listad.TextMatrix(r_int_NumFil - 2, 5) 'CONDICION
         .Cells(r_int_NumFil + 2, 6) = grd_Listad.TextMatrix(r_int_NumFil - 2, 6) 'CUENTA DETRACTORA
         
         .Cells(r_int_NumFil + 2, 7) = grd_Listad.TextMatrix(r_int_NumFil - 2, 7) 'TIPO PERSONAL
         .Cells(r_int_NumFil + 2, 8) = "'" & grd_Listad.TextMatrix(r_int_NumFil - 2, 8) 'CODIGO SICO
         
         r_int_NumFil = r_int_NumFil + 1
      Next r_int_Contad
      
      .Range(.Cells(3, 3), .Cells(3, 8)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_BuscarProv
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "NINGUNA"
   cmb_Buscar.AddItem "NRO DOCUMENTO"
   cmb_Buscar.AddItem "RAZÓN SOCIAL"
   cmb_Buscar.ListIndex = 0
      
   grd_Listad.ColWidth(0) = 0
   grd_Listad.ColWidth(1) = 0
   grd_Listad.ColWidth(2) = 1980
   grd_Listad.ColWidth(3) = 7020
   grd_Listad.ColWidth(4) = 0
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_SetFocus(cmd_Agrega)
End Sub

Public Sub fs_BuscarProv()
   Screen.MousePointer = 11
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAEPRV_TIPDOC, MAEPRV_NUMDOC, MAEPRV_RAZSOC, MAEPRV_CTADET, A.MAEPRV_TIPCNT,  "
   g_str_Parame = g_str_Parame & "        TRIM(B.PARDES_DESCRI) TIPCONTRIB, A.MAEPRV_CONDIC, TRIM(C.PARDES_DESCRI) CONDICION,  "
   g_str_Parame = g_str_Parame & "        MAEPRV_TIPPER, D.PARDES_DESCRI AS TIPOPERSONAL, MAEPRV_CODSIC  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES B ON A.MAEPRV_TIPCNT = B.PARDES_CODITE AND B.PARDES_CODGRP = 119  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES C ON A.MAEPRV_CONDIC = C.PARDES_CODITE AND C.PARDES_CODGRP = 120  "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES D ON A.MAEPRV_TIPPER = D.PARDES_CODITE AND D.PARDES_CODGRP = 127  "
   g_str_Parame = g_str_Parame & "  WHERE MAEPRV_SITUAC = 1  "
   
   If (cmb_Buscar.ListIndex = 1) Then 'numero de documento
       If Len(Trim(txt_Buscar.Text)) > 0 Then
          g_str_Parame = g_str_Parame & "   AND MAEPRV_NUMDOC = '" & Trim(txt_Buscar.Text) & "' "
       End If
   ElseIf (cmb_Buscar.ListIndex = 2) Then 'razon social
       If Len(Trim(txt_Buscar.Text)) > 0 Then
           g_str_Parame = g_str_Parame & "   AND MAEPRV_RAZSOC LIKE '%" & UCase(Trim(txt_Buscar.Text)) & "%'"
       End If
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY MAEPRV_TIPDOC, MAEPRV_RAZSOC ASC  "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún proveedor.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!MAEPRV_TIPDOC)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!maeprv_numdoc & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!MAEPRV_TIPDOC) & "-" & Trim(g_rst_Princi!maeprv_numdoc & "")
      
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!MaePrv_RazSoc & "")
      
      grd_Listad.Col = 4
      grd_Listad.Text = Trim(g_rst_Princi!TIPCONTRIB & "")
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!CONDICION & "")
            
      grd_Listad.Col = 6
      grd_Listad.Text = Trim(g_rst_Princi!MaePrv_CtaDet & "")

      grd_Listad.Col = 7
      grd_Listad.Text = Trim(g_rst_Princi!TIPOPERSONAL & "")
            
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(g_rst_Princi!MAEPRV_CODSIC & "")
      
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(CStr(g_rst_Princi!MAEPRV_TIPDOC)) & Trim(g_rst_Princi!maeprv_numdoc & "") 'ORDEN_GRILLA
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub txt_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call fs_BuscarProv
   Else
      If (cmb_Buscar.ListIndex = 1) Then
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
      Else
          KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
      End If
   End If
End Sub

Private Sub pnl_NumDoc_Click()
   If pnl_NumDoc.Tag = "" Then
      pnl_NumDoc.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 9, "C")
   Else
      pnl_NumDoc.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 9, "C-")
   End If
End Sub

Private Sub pnl_RazSoc_Click()
   If pnl_RazSoc.Tag = "" Then
      pnl_RazSoc.Tag = "1"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_RazSoc.Tag = ""
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub




