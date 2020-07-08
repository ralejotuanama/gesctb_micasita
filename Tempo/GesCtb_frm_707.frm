VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2325
   ClientLeft      =   6825
   ClientTop       =   4515
   ClientWidth     =   5475
   Icon            =   "GesCtb_frm_707.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2415
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5565
      _Version        =   65536
      _ExtentX        =   9816
      _ExtentY        =   4260
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
         TabIndex        =   7
         Top             =   30
         Width           =   5415
         _Version        =   65536
         _ExtentX        =   9551
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
            Height          =   270
            Left            =   630
            TabIndex        =   8
            Top             =   30
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Anexo Nº 5-D"
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
            TabIndex        =   9
            Top             =   270
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   476
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_707.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   750
         Width           =   5415
         _Version        =   65536
         _ExtentX        =   9551
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
            Left            =   630
            Picture         =   "GesCtb_frm_707.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_707.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4800
            Picture         =   "GesCtb_frm_707.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpDet 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_707.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar Detalle"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   3090
            Top             =   90
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
         Height          =   855
         Left            =   30
         TabIndex        =   11
         Top             =   1440
         Width           =   5415
         _Version        =   65536
         _ExtentX        =   9551
         _ExtentY        =   1508
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
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1530
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
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_RepSbs_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_PerMes         As String
Dim l_str_PerAno         As String

Private Sub cmd_ExpDet_Click()
   
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
         
   Call fs_GenDet(l_str_PerMes, l_str_PerAno)
   
End Sub

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
      
   l_str_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   l_str_PerAno = ipp_PerAno.Text
         
   Call fs_GenExc(l_str_PerMes, l_str_PerAno)
      
End Sub

Private Sub cmd_Imprim_Click()

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
      
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   l_str_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   l_str_PerAno = ipp_PerAno.Text
         
   Call fs_GenRep(l_str_PerMes, l_str_PerAno)
   
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_Limpia
   Screen.MousePointer = 0

End Sub

Private Sub fs_Inicia()
         
   cmb_PerMes.Clear
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno = Mid(date, 7, 4)
      
End Sub

Private Sub fs_Limpia()
   cmb_PerMes.ListIndex = -1
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PerMes_Click
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Imprim)
   End If
End Sub

Private Sub cmb_PerMes_Click()
   Call gs_SetFocus(ipp_PerAno)
End Sub

Private Function ff_ValDat(ByVal p_PerMes As String, ByVal p_PerAno As String) As Integer
   
   ff_ValDat = 0
   
   g_str_Parame = "SELECT COUNT(*) AS TOTAL FROM HIS_DEUPRV WHERE "
   g_str_Parame = g_str_Parame & "DEUPRV_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "DEUPRV_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "DEUPRV_NOMREP = 'CTB_REPSBS_02' AND "
   g_str_Parame = g_str_Parame & "DEUPRV_TERCRE = '" & modgen_g_str_NombPC & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      ff_ValDat = g_rst_Princi!Total
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
End Function

Private Sub fs_GenExc(ByVal p_PerMes As String, ByVal p_PerAno As String)
   
   Dim r_obj_Excel            As Excel.Application
   
   Dim r_int_FilCab           As Integer
   Dim r_int_FilDet           As Integer
   Dim r_int_VarAux           As Integer
   Dim r_int_UltDia           As Integer
   Dim r_int_ValDat           As Integer
   Dim r_int_ConAux           As Integer
   
   r_int_UltDia = Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
   r_int_ValDat = ff_ValDat(p_PerMes, p_PerAno)
   
   If r_int_ValDat > 0 Then
      
      g_str_Parame = "DELETE FROM HIS_DEUPRV WHERE "
      g_str_Parame = g_str_Parame & "DEUPRV_PERMES = " & l_str_PerMes & " AND "
      g_str_Parame = g_str_Parame & "DEUPRV_PERANO = " & l_str_PerAno & " AND "
      g_str_Parame = g_str_Parame & "DEUPRV_NOMREP = 'CTB_REPSBS_02' AND "
      g_str_Parame = g_str_Parame & "DEUPRV_TERCRE = '" & modgen_g_str_NombPC & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
        
   End If
      
   Call fs_CalDat
      
   'llamada de Tabla para la Exportacion de Datos
   g_str_Parame = "SELECT * FROM HIS_DEUPRV WHERE "
   g_str_Parame = g_str_Parame & "DEUPRV_PERMES = '" & l_str_PerMes & "' AND "
   g_str_Parame = g_str_Parame & "DEUPRV_PERANO = '" & l_str_PerAno & "' AND "
   g_str_Parame = g_str_Parame & "DEUPRV_NOMREP = 'CTB_REPSBS_02' AND "
   g_str_Parame = g_str_Parame & "DEUPRV_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "ORDER BY DEUPRV_GARLIN, DEUPRV_SUBCAB ASC"
          
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Operaciones registradas.", vbInformation, modgen_g_str_NomPlt
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
      .Cells(10, 2) = "Al " & r_int_UltDia & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & l_str_PerAno
      .Cells(11, 2) = "(En nuevos soles)"
      .Cells(13, 2) = "INFORME DE CLASIFICACIÓN DE LOS DEUDORES DE LA CARTERA DE CRÉDITOS, CONTINGENTES Y ARRENDAMIENTOS"
      .Cells(14, 2) = "FINANCIEROS QUE RESPALDAN FINANCIAMIENTOS O LINEAS DE CREDITO 1/"
      
      .Cells(166, 3) = "GERENTE"
      .Cells(166, 5) = "CONTADOR"
      .Cells(166, 7) = "FUNCIONARIO"
           
      .Cells(167, 3) = "GENERAL"
      .Cells(167, 5) = "GENERAL"
      .Cells(167, 7) = "RESPONSABLE"
      
      .Cells(166, 3).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(166, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(167, 3).HorizontalAlignment = xlHAlignCenter
      
      .Cells(166, 5).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(166, 5).HorizontalAlignment = xlHAlignCenter
      .Cells(167, 5).HorizontalAlignment = xlHAlignCenter
      
      .Cells(166, 7).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(166, 7).HorizontalAlignment = xlHAlignCenter
      .Cells(167, 7).HorizontalAlignment = xlHAlignCenter
      
      .Cells(16, 2) = "Línea Mivivienda:"
      .Cells(16, 2).Font.Bold = True
      
      .Cells(65, 2) = "Línea BID:"
      .Cells(65, 2).Font.Bold = True
      
      .Cells(114, 2) = "Línea Fidel Ramirez Prado:"
      .Cells(114, 2).Font.Bold = True
                       
      For r_int_FilCab = 18 To 154
                  
         If r_int_FilCab = 54 Then
         
            .Cells(r_int_FilCab, 2) = "TOTAL"
            r_int_FilCab = r_int_FilCab + 1
         
         ElseIf r_int_FilCab = 103 Then
            
            .Cells(r_int_FilCab, 2) = "TOTAL"
            r_int_FilCab = r_int_FilCab + 1
         
         ElseIf r_int_FilCab = 152 Then
            
            .Cells(r_int_FilCab, 2) = "TOTAL"
            r_int_FilCab = r_int_FilCab + 1
         
         ElseIf r_int_FilCab = 64 Or r_int_FilCab = 113 Then
            
            r_int_FilCab = r_int_FilCab + 3
         
         End If
                  
         If r_int_FilCab = 18 Or r_int_FilCab = 67 Or r_int_FilCab = 116 Then
             
            .Cells(r_int_FilCab, 2) = "A.- MONTO DE LOS CRÉDITOS, CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS 2/"
            .Range(.Cells(r_int_FilCab, 3), .Cells(r_int_FilCab, 8)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).VerticalAlignment = xlCenter
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Font.Bold = True
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Borders.Color = RGB(0, 0, 0)
             
         ElseIf r_int_FilCab = 27 Or r_int_FilCab = 76 Or r_int_FilCab = 125 Then
            
            .Cells(r_int_FilCab, 2) = "B.- PROVISIONES CONSTITUIDAS 3/"
            .Range(.Cells(r_int_FilCab, 3), .Cells(r_int_FilCab, 8)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).VerticalAlignment = xlCenter
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Font.Bold = True
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Borders.Color = RGB(0, 0, 0)
         
         ElseIf r_int_FilCab = 36 Or r_int_FilCab = 85 Or r_int_FilCab = 134 Then
            
            .Cells(r_int_FilCab, 2) = "C.- PROVISIONES REQUERIDAS 4/"
            .Range(.Cells(r_int_FilCab, 3), .Cells(r_int_FilCab, 8)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).VerticalAlignment = xlCenter
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Font.Bold = True
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Borders.Color = RGB(0, 0, 0)
         
         ElseIf r_int_FilCab = 45 Or r_int_FilCab = 94 Or r_int_FilCab = 143 Then
            
            .Cells(r_int_FilCab, 2) = "D.- SUPERÁVIT (DÉFICIT) DE PROVISIONES 5/"
            .Range(.Cells(r_int_FilCab, 3), .Cells(r_int_FilCab, 8)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).VerticalAlignment = xlCenter
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Font.Bold = True
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Borders.Color = RGB(0, 0, 0)
            
            .Range(.Cells(r_int_FilCab + 9, 2), .Cells(r_int_FilCab + 9, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(r_int_FilCab + 9, 2), .Cells(r_int_FilCab + 9, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(r_int_FilCab + 9, 2), .Cells(r_int_FilCab + 9, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(r_int_FilCab + 9, 2), .Cells(r_int_FilCab + 9, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
         
         ElseIf r_int_FilCab = 55 Or r_int_FilCab = 104 Or r_int_FilCab = 153 Then
            
            .Cells(r_int_FilCab, 2) = "E.- MONTO DE LOS CREDITOS, CONTINGENTES 6/"
            .Range(.Cells(r_int_FilCab, 3), .Cells(r_int_FilCab, 8)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).VerticalAlignment = xlCenter
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Font.Bold = True
            .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Borders.Color = RGB(0, 0, 0)
            
            .Range(.Cells(r_int_FilCab + 8, 2), .Cells(r_int_FilCab + 8, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         
         End If
         
         .Cells(r_int_FilCab, 3) = "Normal"
         .Cells(r_int_FilCab, 4) = "CPP"
         .Cells(r_int_FilCab, 5) = "Deficiente"
         .Cells(r_int_FilCab, 6) = "Dudoso"
         .Cells(r_int_FilCab, 7) = "Pérdida"
         .Cells(r_int_FilCab, 8) = "Total"
                 
         r_int_FilCab = r_int_FilCab + 8
                 
      Next r_int_FilCab
              
      For r_int_FilDet = 19 To 160
         
         If r_int_FilDet = 55 Or r_int_FilDet = 104 Or r_int_FilDet = 153 Then
            r_int_FilDet = r_int_FilDet + 1
         End If
         
         .Cells(r_int_FilDet + 0, 2) = "Corporativos"
         .Cells(r_int_FilDet + 1, 2) = "Grandes Empresas"
         .Cells(r_int_FilDet + 2, 2) = "Medianas Empresas"
         .Cells(r_int_FilDet + 3, 2) = "Pequeñas Empresas"
         .Cells(r_int_FilDet + 4, 2) = "Microempresas"
         .Cells(r_int_FilDet + 5, 2) = "Consumo revolvente"
         .Cells(r_int_FilDet + 6, 2) = "Consumo no revolvente"
         .Cells(r_int_FilDet + 7, 2) = "Hipotecario para Vivienda"
         
         For r_int_VarAux = 0 To 7
         
            .Range(.Cells(r_int_FilDet + r_int_VarAux, 2), .Cells(r_int_FilDet + r_int_VarAux, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(r_int_FilDet + r_int_VarAux, 2), .Cells(r_int_FilDet + r_int_VarAux, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(r_int_FilDet + r_int_VarAux, 2), .Cells(r_int_FilDet + r_int_VarAux, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
         
         Next r_int_VarAux
                   
         If r_int_FilDet = 56 Or r_int_FilDet = 105 Then
            r_int_FilDet = r_int_FilDet + 11
         Else
            r_int_FilDet = r_int_FilDet + 8
         End If
                 
      Next r_int_FilDet
        
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
      .Columns("I").ColumnWidth = 2
      
      .Range(.Cells(1, 1), .Cells(400, 400)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(400, 400)).Font.Size = 10
                    
   End With
   
   g_rst_Princi.MoveFirst
        
   Do While Not g_rst_Princi.EOF
        
      'Mi Vivienda
      If Trim(g_rst_Princi!DEUPRV_GARLIN) = "000001" Then
      
         If Trim(g_rst_Princi!DEUPRV_SUBCAB) = "A" Then
            r_int_ConAux = 26
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "B" Then
            r_int_ConAux = 35
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "C" Then
            r_int_ConAux = 44
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "D" Then
            r_int_ConAux = 53
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "E" Then
            r_int_ConAux = 62
         End If
      
      'BID
      ElseIf g_rst_Princi!DEUPRV_GARLIN = "000002" Then
      
         If Trim(g_rst_Princi!DEUPRV_SUBCAB) = "A" Then
            r_int_ConAux = 75
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "B" Then
            r_int_ConAux = 84
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "C" Then
            r_int_ConAux = 93
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "D" Then
            r_int_ConAux = 102
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "E" Then
            r_int_ConAux = 111
         End If
         
      'Fidel Ramirez
      ElseIf g_rst_Princi!DEUPRV_GARLIN = "000003" Then
         
         If Trim(g_rst_Princi!DEUPRV_SUBCAB) = "A" Then
            r_int_ConAux = 124
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "B" Then
            r_int_ConAux = 133
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "C" Then
            r_int_ConAux = 142
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "D" Then
            r_int_ConAux = 151
         ElseIf Trim(g_rst_Princi!DEUPRV_SUBCAB) = "E" Then
            r_int_ConAux = 160
         End If
      
      End If
      
      If r_int_ConAux = 62 Or r_int_ConAux = 111 Or r_int_ConAux = 160 Then
         r_int_ConAux = r_int_ConAux + 1
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 3) = Format(g_rst_Princi!DEUPRV_MTONOR, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 4) = Format(g_rst_Princi!DEUPRV_MTOCPP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 5) = Format(g_rst_Princi!DEUPRV_MTODEF, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 6) = Format(g_rst_Princi!DEUPRV_MTODUD, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 7) = Format(g_rst_Princi!DEUPRV_MTOPER, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConAux, 8) = Format(g_rst_Princi!DEUPRV_MTONOR + g_rst_Princi!DEUPRV_MTOCPP + g_rst_Princi!DEUPRV_MTODEF + g_rst_Princi!DEUPRV_MTODUD + g_rst_Princi!DEUPRV_MTOPER, "###,###,##0.00")
      
      
      r_obj_Excel.ActiveSheet.Cells(54, 3) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(54, 4) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(54, 5) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(54, 6) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(54, 7) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(54, 8) = Format(0, "###,###,##0.00")
      
      r_obj_Excel.ActiveSheet.Cells(103, 3) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(103, 4) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(103, 5) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(103, 6) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(103, 7) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(103, 8) = Format(0, "###,###,##0.00")
      
      r_obj_Excel.ActiveSheet.Cells(152, 3) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(152, 4) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(152, 5) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(152, 6) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(152, 7) = Format(0, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(152, 8) = Format(0, "###,###,##0.00")

      g_rst_Princi.MoveNext
      DoEvents
   
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
  
   Screen.MousePointer = 0
  
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing

End Sub

Private Sub fs_GenDet(ByVal p_PerMes As String, ByVal p_PerAno As String)

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
      
   g_str_Parame = "SELECT HIPCIE_NUMOPE, HIPMAE_GARLIN, HIPMAE_CODPRD, HIPCIE_SITCRE, HIPCIE_CLACRE, HIPCIE_CLAPRV, HIPCIE_SALCAP, HIPCIE_SALCON, HIPCIE_PRVGEN, HIPCIE_PRVESP, HIPCIE_PRVCIC, HIPCIE_TIPCAM FROM CRE_HIPCIE H, CRE_HIPMAE M WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_NUMOPE = HIPMAE_NUMOPE AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPMAE_GARLIN, HIPCIE_NUMOPE, HIPCIE_SITCRE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Operaciones registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM."
      .Cells(1, 2) = "NRO OPERACION."
      .Cells(1, 3) = "LINEA DE GARANTIA."
      .Cells(1, 4) = "TIPO DE PRODUCTO."
      .Cells(1, 5) = "SITUACION DEL CREDITO."
      .Cells(1, 6) = "CLASIF. DEL CREDITO."
      .Cells(1, 7) = "CLASIF. PROVISION."
      .Cells(1, 8) = "TIPO DE CAMBIO."
      .Cells(1, 9) = "SALDO CAPITAL MON.ORG."
      .Cells(1, 10) = "SALDO CONCESIONAL MON.ORG."
      .Cells(1, 11) = "PROV. GENERICA EN SOLES"
      .Cells(1, 12) = "PROV. ESPECIFICA EN SOLES."
      .Cells(1, 13) = "PROV. PRO-CICLICA EN SOLES."
         
      .Range(.Cells(1, 1), .Cells(1, 13)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 13)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 5
      
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      
      .Columns("C").ColumnWidth = 28
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 28
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 22
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 19
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 18
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      
      .Columns("H").ColumnWidth = 15
      .Columns("I").ColumnWidth = 30
      .Columns("J").ColumnWidth = 30
      .Columns("K").ColumnWidth = 30
      .Columns("L").ColumnWidth = 30
      .Columns("M").ColumnWidth = 30
                 
   End With
   
   g_rst_Princi.MoveFirst
     
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
         
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!HIPCIE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = moddat_gf_Consulta_ParDes("306", g_rst_Princi!HIPMAE_GARLIN)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = g_rst_Princi!HIPCIE_SITCRE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = g_rst_Princi!HIPCIE_CLACRE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = g_rst_Princi!HIPCIE_CLAPRV
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Format(g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!HIPCIE_SALCAP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Format(g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!HIPCIE_PRVGEN, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!HIPCIE_PRVESP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!HIPCIE_PRVCIC, "###,###,##0.00")
         
      r_int_ConVer = r_int_ConVer + 1
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_CalDat()
        
   Dim r_int_Contad           As Integer
   Dim r_int_ConAux           As Integer
   
   Dim r_dbl_MtoNor(14)       As Double
   Dim r_dbl_MtoCpp(14)       As Double
   Dim r_dbl_MtoDef(14)       As Double
   Dim r_dbl_MtoDud(14)       As Double
   Dim r_dbl_MtoPer(14)       As Double
   
   Dim r_str_CabRep(14)       As String
   Dim r_str_Garlin(14)       As String
   
   r_int_Contad = 0
         
   'Leer Tabla de Creditos del mes CRE_HIPCIE
   g_str_Parame = "SELECT HIPCIE_NUMOPE, HIPCIE_CLAPRV, HIPCIE_PRVGEN, HIPCIE_PRVESP, HIPCIE_PRVCIC, HIPCIE_SITCRE, HIPCIE_CLACRE, HIPCIE_CLAPRV, HIPCIE_SALCAP, HIPCIE_SALCON, HIPMAE_GARLIN, HIPMAE_CODPRD, HIPCIE_TIPCAM FROM CRE_HIPCIE H, CRE_HIPMAE M WHERE "
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
   
   
   For r_int_ConAux = 0 To UBound(r_str_CabRep)
   
      r_str_CabRep(0 + r_int_ConAux) = "A"
      r_str_CabRep(1 + r_int_ConAux) = "B"
      r_str_CabRep(2 + r_int_ConAux) = "C"
      r_str_CabRep(3 + r_int_ConAux) = "D"
      r_str_CabRep(4 + r_int_ConAux) = "E"
      
      r_int_ConAux = r_int_ConAux + 4
      
   Next r_int_ConAux
   
      
   g_rst_Princi.MoveFirst
     
   Do While Not g_rst_Princi.EOF
      
      If g_rst_Princi!HIPMAE_GARLIN = "000001" Then
         
         r_str_Garlin(0) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(1) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(2) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(3) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(4) = g_rst_Princi!HIPMAE_GARLIN
         
         If g_rst_Princi!HIPMAE_CODPRD = "003" Or g_rst_Princi!HIPMAE_CODPRD = "004" Or g_rst_Princi!HIPMAE_CODPRD = "007" Then
            'Vigente
            If g_rst_Princi!HIPCIE_SITCRE = 1 Then
               'Normal
               If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
                  r_dbl_MtoNor(0) = r_dbl_MtoNor(0) + g_rst_Princi!HIPCIE_SALCAP
                  r_dbl_MtoNor(1) = r_dbl_MtoNor(1) + (g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVCIC)
                  'r_dbl_MtoNor(2) = r_dbl_MtoNor(2) + (g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVCIC)
               End If
            'Vencido
            ElseIf g_rst_Princi!HIPCIE_SITCRE = 5 Then
               'CPP
               If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
                  r_dbl_MtoCpp(0) = r_dbl_MtoCpp(0) + g_rst_Princi!HIPCIE_SALCAP
                  r_dbl_MtoCpp(1) = r_dbl_MtoCpp(1) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoCpp(2) = r_dbl_MtoCpp(2) + g_rst_Princi!HIPCIE_PRVESP
               'Deficiente
               ElseIf g_rst_Princi!HIPCIE_CLAPRV = 2 Then
                  r_dbl_MtoDef(0) = r_dbl_MtoDef(0) + g_rst_Princi!HIPCIE_SALCAP
                  r_dbl_MtoDef(1) = r_dbl_MtoDef(1) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoDef(2) = r_dbl_MtoDef(2) + g_rst_Princi!HIPCIE_PRVESP
               'Dudoso
               ElseIf g_rst_Princi!HIPCIE_CLAPRV = 3 Then
                  r_dbl_MtoDud(0) = r_dbl_MtoDud(0) + g_rst_Princi!HIPCIE_SALCAP
                  r_dbl_MtoDud(1) = r_dbl_MtoDud(1) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoDud(2) = r_dbl_MtoDud(2) + g_rst_Princi!HIPCIE_PRVESP
               'Perdida
               ElseIf g_rst_Princi!HIPCIE_CLAPRV = 4 Then
                  r_dbl_MtoPer(0) = r_dbl_MtoPer(0) + g_rst_Princi!HIPCIE_SALCAP
                  r_dbl_MtoPer(1) = r_dbl_MtoPer(1) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoPer(2) = r_dbl_MtoPer(2) + g_rst_Princi!HIPCIE_PRVESP
               End If
            End If
         End If
      ElseIf g_rst_Princi!HIPMAE_GARLIN = "000002" Then
      
         r_str_Garlin(5) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(6) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(7) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(8) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(9) = g_rst_Princi!HIPMAE_GARLIN
  
         If g_rst_Princi!HIPMAE_CODPRD = "001" Or g_rst_Princi!HIPMAE_CODPRD = "002" Then
            If g_rst_Princi!HIPCIE_SITCRE = 1 Then
               If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
                  r_dbl_MtoNor(5) = r_dbl_MtoNor(5) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_dbl_MtoNor(6) = r_dbl_MtoNor(6) + (g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVCIC)
                  'r_dbl_MtoNor(7) = r_dbl_MtoNor(7) + (g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVCIC)
               End If
            ElseIf g_rst_Princi!HIPCIE_SITCRE = 5 Then
               If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
                  r_dbl_MtoCpp(5) = r_dbl_MtoCpp(5) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_dbl_MtoCpp(6) = r_dbl_MtoCpp(6) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoCpp(7) = r_dbl_MtoCpp(7) + g_rst_Princi!HIPCIE_PRVESP
               ElseIf g_rst_Princi!HIPCIE_CLAPRV = 2 Then
                  r_dbl_MtoDef(5) = r_dbl_MtoDef(5) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_dbl_MtoDef(6) = r_dbl_MtoDef(6) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoDef(7) = r_dbl_MtoDef(7) + g_rst_Princi!HIPCIE_PRVESP
               ElseIf g_rst_Princi!HIPCIE_CLAPRV = 3 Then
                  r_dbl_MtoDud(5) = r_dbl_MtoDud(5) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_dbl_MtoDud(6) = r_dbl_MtoDud(6) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoDud(7) = r_dbl_MtoDud(7) + g_rst_Princi!HIPCIE_PRVESP
               ElseIf g_rst_Princi!HIPCIE_CLAPRV = 4 Then
                  r_dbl_MtoPer(5) = r_dbl_MtoPer(5) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_dbl_MtoPer(6) = r_dbl_MtoPer(6) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoPer(7) = r_dbl_MtoPer(7) + g_rst_Princi!HIPCIE_PRVESP
               End If
            End If
         End If
  
      ElseIf g_rst_Princi!HIPMAE_GARLIN = "000003" Then
      
         r_str_Garlin(10) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(11) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(12) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(13) = g_rst_Princi!HIPMAE_GARLIN
         r_str_Garlin(14) = g_rst_Princi!HIPMAE_GARLIN
  
         If g_rst_Princi!HIPMAE_CODPRD = "002" Then
            If g_rst_Princi!HIPCIE_SITCRE = 1 Then
               If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
                  r_dbl_MtoNor(10) = r_dbl_MtoNor(10) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_dbl_MtoNor(11) = r_dbl_MtoNor(11) + (g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVCIC)
                  'r_dbl_MtoNor(12) = r_dbl_MtoNor(12) + (g_rst_Princi!HIPCIE_PRVGEN + g_rst_Princi!HIPCIE_PRVCIC)
               End If
            ElseIf g_rst_Princi!HIPCIE_SITCRE = 5 Then
               If g_rst_Princi!HIPCIE_CLAPRV = 1 Then
                  r_dbl_MtoCpp(10) = r_dbl_MtoCpp(10) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_dbl_MtoCpp(11) = r_dbl_MtoCpp(11) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoCpp(12) = r_dbl_MtoCpp(12) + g_rst_Princi!HIPCIE_PRVESP
               ElseIf g_rst_Princi!HIPCIE_CLAPRV = 2 Then
                  r_dbl_MtoDef(10) = r_dbl_MtoDef(10) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_dbl_MtoDef(11) = r_dbl_MtoDef(11) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoDef(12) = r_dbl_MtoDef(12) + g_rst_Princi!HIPCIE_PRVESP
               ElseIf g_rst_Princi!HIPCIE_CLAPRV = 3 Then
                  r_dbl_MtoDud(10) = r_dbl_MtoDud(10) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_dbl_MtoDud(11) = r_dbl_MtoDud(11) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoDud(12) = r_dbl_MtoDud(12) + g_rst_Princi!HIPCIE_PRVESP
               ElseIf g_rst_Princi!HIPCIE_CLAPRV = 4 Then
                  r_dbl_MtoPer(10) = r_dbl_MtoPer(10) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
                  r_dbl_MtoPer(11) = r_dbl_MtoPer(11) + g_rst_Princi!HIPCIE_PRVESP
                  'r_dbl_MtoPer(12) = r_dbl_MtoPer(12) + g_rst_Princi!HIPCIE_PRVESP
               End If
            End If
         End If
      End If
  
      g_rst_Princi.MoveNext
      DoEvents
  
   Loop
   
   r_dbl_MtoNor(2) = r_dbl_MtoNor(1)
   r_dbl_MtoCpp(2) = r_dbl_MtoCpp(1)
   r_dbl_MtoDef(2) = r_dbl_MtoDef(1)
   r_dbl_MtoDud(2) = r_dbl_MtoDud(1)
   r_dbl_MtoPer(2) = r_dbl_MtoPer(1)
   
   r_dbl_MtoNor(7) = r_dbl_MtoNor(6)
   r_dbl_MtoCpp(7) = r_dbl_MtoCpp(6)
   r_dbl_MtoDef(7) = r_dbl_MtoDef(6)
   r_dbl_MtoDud(7) = r_dbl_MtoDud(6)
   r_dbl_MtoPer(7) = r_dbl_MtoPer(6)
   
   r_dbl_MtoNor(12) = r_dbl_MtoNor(11)
   r_dbl_MtoCpp(12) = r_dbl_MtoCpp(11)
   r_dbl_MtoDef(12) = r_dbl_MtoDef(11)
   r_dbl_MtoDud(12) = r_dbl_MtoDud(11)
   r_dbl_MtoPer(12) = r_dbl_MtoPer(11)
    
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
    
   'Insertando a la Tabla
   For r_int_Contad = 0 To UBound(r_str_CabRep)
      g_str_Parame = "USP_HIS_DEUPRV ("
      g_str_Parame = g_str_Parame & "'CTB_REPSBS_02', "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(r_str_Garlin(r_int_Contad)) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(r_str_CabRep(r_int_Contad)) & "', "
      g_str_Parame = g_str_Parame & 13 & ", "
      g_str_Parame = g_str_Parame & CDbl(r_dbl_MtoNor(r_int_Contad)) & ", "
      g_str_Parame = g_str_Parame & CDbl(r_dbl_MtoCpp(r_int_Contad)) & ", "
      g_str_Parame = g_str_Parame & CDbl(r_dbl_MtoDef(r_int_Contad)) & ", "
      g_str_Parame = g_str_Parame & CDbl(r_dbl_MtoDud(r_int_Contad)) & ", "
      g_str_Parame = g_str_Parame & CDbl(r_dbl_MtoPer(r_int_Contad)) & ")"
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_CTB_SALPRV.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
   Next r_int_Contad
      
End Sub

Private Sub fs_GenRep(ByVal p_PerMes As String, ByVal p_PerAno As String)

   Dim r_obj_Excel            As Excel.Application
   Dim r_int_UltDia           As Integer
   Dim r_int_ValDat           As Integer
      
   r_int_UltDia = Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
   r_int_ValDat = ff_ValDat(p_PerMes, p_PerAno)
   
   If r_int_ValDat > 0 Then
      
      g_str_Parame = "DELETE FROM HIS_DEUPRV WHERE "
      g_str_Parame = g_str_Parame & "DEUPRV_PERMES = " & l_str_PerMes & " AND "
      g_str_Parame = g_str_Parame & "DEUPRV_PERANO = " & l_str_PerAno & " AND "
      g_str_Parame = g_str_Parame & "DEUPRV_NOMREP = 'CTB_REPSBS_02' AND "
      g_str_Parame = g_str_Parame & "DEUPRV_TERCRE = '" & modgen_g_str_NombPC & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
        
   End If

   Call fs_CalDat
              
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
    
   crp_Imprim.DataFiles(0) = UCase(moddat_g_str_EntDat) & ".HIS_DEUPRV"
    
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_20.RPT"
           
   crp_Imprim.SelectionFormula = "{HIS_DEUPRV.DEUPRV_PERMES} = " & p_PerMes & " AND {HIS_DEUPRV.DEUPRV_PERANO} = " & p_PerAno & " AND {HIS_DEUPRV.DEUPRV_USUCRE} = '" & modgen_g_str_CodUsu & "' "
   
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   Screen.MousePointer = 0
  
End Sub
