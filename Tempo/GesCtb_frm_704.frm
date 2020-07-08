VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   2580
   ClientLeft      =   7665
   ClientTop       =   4245
   ClientWidth     =   6195
   Icon            =   "GesCtb_frm_704.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2625
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6225
      _Version        =   65536
      _ExtentX        =   10980
      _ExtentY        =   4630
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
         Top             =   60
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
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
            Caption         =   "Anexo Nº 5"
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
            Width           =   5325
            _Version        =   65536
            _ExtentX        =   9393
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0102-01 Informe de Clasificación de Deudores y Provisiones"
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
            Picture         =   "GesCtb_frm_704.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   780
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
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
         Begin VB.CommandButton cmd_ExpArc 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_704.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_704.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_704.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5520
            Picture         =   "GesCtb_frm_704.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpDet 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_704.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir Reporte"
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
         Height          =   1095
         Left            =   30
         TabIndex        =   11
         Top             =   1470
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
         _ExtentY        =   1931
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
         Begin VB.CheckBox chk_PrvPro 
            Caption         =   "Incluir Provisión Prociclica"
            Height          =   285
            Left            =   1530
            TabIndex        =   15
            Top             =   780
            Width           =   2445
         End
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
            TabIndex        =   14
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
Attribute VB_Name = "frm_RepSbs_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_Evalua(773)      As Double
Dim l_str_PerMes           As String
Dim l_str_PerAno           As String
Dim l_int_PrvPro           As Integer

Private Sub cmd_ExpArc_Click()
   
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
   
   Call fs_GenArc(l_str_PerMes, l_str_PerAno)
      
End Sub

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
   
   Call fs_GenExc_Det
   
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
   l_int_PrvPro = chk_PrvPro.Value
         
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

Private Sub fs_GenArc(ByVal p_PerMes As String, ByVal p_PerAno As String)
   
   Dim r_int_NumRes     As Integer
   Dim r_int_PerMes     As Integer
   Dim r_int_PerAno     As Integer
   Dim r_int_Contad     As Integer
   Dim r_int_ConGen     As Integer
   Dim r_int_ConTem     As Integer
   Dim r_int_CodEmp     As Integer
   Dim r_int_ConAux     As Integer
   Dim r_int_ValDat     As Integer
   
   Dim r_str_Cadena     As String
   Dim r_str_NomRes     As String
   Dim r_str_FecRpt     As String
   
   Screen.MousePointer = 11
      
   r_int_ValDat = ff_ValDat(p_PerMes, p_PerAno)
   
   If r_int_ValDat > 0 Then
   
      g_str_Parame = "DELETE FROM HIS_CLADEU WHERE "
      g_str_Parame = g_str_Parame & "CLADEU_PERMES = " & l_str_PerMes & " AND "
      g_str_Parame = g_str_Parame & "CLADEU_PERANO = " & l_str_PerAno & " "
      'g_str_Parame = g_str_Parame & "CLADEU_NOMREP = 'CTB_REPSBS_05' AND "
      'g_str_Parame = g_str_Parame & "CLADEU_TERCRE = '" & modgen_g_str_NombPC & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
        
   End If
   
   Call fs_CalDat
   
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
   
   r_str_NomRes = "C:\01" & Right(r_int_PerAno, 2) & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".105"
   
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
   
   g_str_Parame = "SELECT * FROM MNT_EMPGRP "
   g_str_Parame = g_str_Parame & "WHERE EMPGRP_SITUAC = 1"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_int_CodEmp = g_rst_Princi!EMPGRP_CODSBS
      
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_int_ConTem = 0
   
   Print #r_int_NumRes, Format(105, "0000") & Format(1, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
      
   For r_int_ConGen = 100 To 12300 Step 100
      r_str_Cadena = ""
            
      If r_int_ConGen <> 100 And r_int_ConGen <> 1100 And r_int_ConGen <> 2100 And r_int_ConGen <> 3100 And r_int_ConGen <> 4100 And r_int_ConGen <> 5200 And r_int_ConGen <> 6200 And r_int_ConGen <> 6400 And r_int_ConGen <> 7200 And r_int_ConGen <> 8200 And r_int_ConGen <> 8400 And r_int_ConGen <> 9400 And r_int_ConGen <> 10400 And r_int_ConGen <> 11400 Then
            
         For r_int_ConAux = 0 To 5 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(l_str_Evalua(r_int_ConTem), "########0.00"), 1, "0", 18)
            r_int_ConTem = r_int_ConTem + 1
         Next
      Else
         For r_int_ConAux = 0 To 5 Step 1
            r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 18)
         Next
      
      End If
          
      Print #r_int_NumRes, Format(r_int_ConGen, "000000") & r_str_Cadena
      
      If r_int_ConGen = 4700 Then
         r_int_ConGen = r_int_ConGen + 100
      End If
      
   Next
         
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   
   Screen.MousePointer = 0
   
   MsgBox "Archivo creado.", vbInformation, modgen_g_str_NomPlt
   
   
End Sub


Private Sub fs_GenExc(ByVal p_PerMes As String, ByVal p_PerAno As String)
   
   Dim r_obj_Excel            As Excel.Application
   
   Dim r_int_FilCab           As Integer
   Dim r_int_FilDet           As Integer
   Dim r_int_VarAux           As Integer
   Dim r_int_UltDia           As Integer
   Dim r_int_ValDat           As Integer
   Dim r_int_ConAux           As Integer
   Dim r_int_ColExc           As Integer
   Dim r_int_FilExc           As Integer
   Dim r_int_ValArr           As Integer
   
   Screen.MousePointer = 11
      
   r_int_UltDia = Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
   r_int_ValDat = ff_ValDat(p_PerMes, p_PerAno)
   
   If r_int_ValDat > 0 Then
   
      g_str_Parame = "DELETE FROM HIS_CLADEU WHERE "
      g_str_Parame = g_str_Parame & "CLADEU_PERMES = " & l_str_PerMes & " AND "
      g_str_Parame = g_str_Parame & "CLADEU_PERANO = " & l_str_PerAno & " "
      'g_str_Parame = g_str_Parame & "CLADEU_NOMREP = 'CTB_REPSBS_05' AND "
      'g_str_Parame = g_str_Parame & "CLADEU_TERCRE = '" & modgen_g_str_NombPC & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
        
   End If
   
   
      
   Call fs_CalDat
      
   'llamada de Tabla para la Exportacion de Datos
   g_str_Parame = "SELECT * FROM HIS_CLADEU WHERE "
   g_str_Parame = g_str_Parame & "CLADEU_PERMES = '" & l_str_PerMes & "' AND "
   g_str_Parame = g_str_Parame & "CLADEU_PERANO = '" & l_str_PerAno & "' "
   'g_str_Parame = g_str_Parame & "CLADEU_NOMREP = 'CTB_REPSBS_05' AND "
   'g_str_Parame = g_str_Parame & "CLADEU_TERCRE = '" & modgen_g_str_NombPC & "' "
   g_str_Parame = g_str_Parame & "ORDER BY CLADEU_SUBCAB ASC"
          
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
      
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se encontraron Operaciones registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      
      .Cells(5, 2) = "ANEXO Nº 5"
      .Cells(6, 2) = "INFORME DE CLASIFICACION DE DEUDORES Y PROVISIONES"
      .Cells(8, 2) = "EDPYME MICASITA"
      .Cells(8, 8) = "CODIGO: 00240"
      .Cells(10, 2) = "Al " & r_int_UltDia & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & l_str_PerAno
      .Cells(11, 2) = "(En nuevos soles)"
      .Cells(13, 2) = "INFORME DE CLASIFICACIÓN DE LOS DEUDORES DE LA CARTERA DE CRÉDITOS DIRECTOS E INDIRECTOS"
      .Cells(142, 2) = "CUADRE DEL ANEXO N° 5 CON CIFRAS DEL BALANCE 22/"
      
      .Cells(145, 2) = "V.- CIFRAS DEL BALANCE"
      .Cells(147, 2) = "CREDITOS DIRECTOS"
      .Cells(148, 2) = "Créditos directos: 1401+1403+1404+1405+1406-2901.01-2901.02-2901.04"
      .Cells(149, 2) = "CREDITOS INDIRECTOS"
      .Cells(150, 2) = "a) Confirmaciones de cartas de crédito irrevocables, de hasta un año, cuando el banco emisor sea una empresa del sistema financiero del exterior de primer nivel"
      .Cells(151, 2) = "b) Emisiones de cartas fianzas que respalden obligaciones de hacer y no hacer"
      .Cells(152, 2) = "c) Emisiones de avales, cartas de crédito de importación y cartas fianzas no incluidas en el literal 'b)', y las confirmaciones de cartas de crédito no incluidas en el literal 'a)' asi como las aceptaciones bancarias"
      .Cells(153, 2) = "Total"
      .Cells(154, 2) = "W.- ANEXO 5"
      .Cells(156, 2) = "Total"
      .Cells(146, 3) = "Saldo"
      .Cells(146, 4) = "Exposicion equivalente a riesgo crediticio"
      .Cells(146, 5) = "Provisiones Genéricas"
      .Cells(146, 6) = "Provisiones Especificas"
      
      .Cells(155, 4) = "Creditos Directos e Indirectos Afectos a Provisiones"
      .Cells(155, 5) = "Provisiones Genéricas Constituidas"
      .Cells(155, 6) = "Provisiones Especificas Constituidas"
      
      .Cells(161, 2) = "____________________________"
      .Cells(161, 4) = "____________________________"
      .Cells(161, 8) = "____________________________"
      
      .Cells(162, 2) = "Sr. Roberto Baba Yamamoto"
      .Cells(162, 4) = "Srta. Rossana Mesa Bustamente"
      .Cells(162, 8) = "Sr. Javier Delgado Blanco"
           
      .Cells(163, 2) = "Gerente General"
      .Cells(163, 4) = "Contador General"
      .Cells(163, 8) = "Unidad de Riesgos"
            
      .Range(.Cells(161, 2), .Cells(163, 8)).HorizontalAlignment = xlHAlignCenter
            
      .Cells(77, 2) = "Consumo no revolvente"
      .Cells(97, 2) = "Con cobertura del Fondo MiVivienda 11/"
      
      .Range(.Cells(77, 2), .Cells(77, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(77, 2), .Cells(77, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(77, 2), .Cells(77, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(99, 2), .Cells(99, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(99, 2), .Cells(99, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(99, 2), .Cells(99, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
                       
      For r_int_FilCab = 16 To 137
         If r_int_FilCab = 16 Then
            .Cells(r_int_FilCab, 2) = "A.- MONTO DE LOS CRÉDITOS DIRECTOS E INDIRECTOS 1/"
         ElseIf r_int_FilCab = 26 Then
            .Cells(r_int_FilCab, 2) = "A'.- MONTO DE LOS CREDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO CREDITICIO DE LOS CREDITOS INDIRECTOS 2/"
         ElseIf r_int_FilCab = 36 Then
            .Cells(r_int_FilCab, 2) = "B.- NUMERO DE DEUDORES 3/"
         ElseIf r_int_FilCab = 46 Then
            .Cells(r_int_FilCab, 2) = "C.- MONTO DE LOS CREDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO CREDITICIO DE LOS CREDITOS INDIRECTOS CON SUSTITUCION DE CONTRAPARTE CREDITICIA - ANTES DE LA SUSTITUCION 5/"
         ElseIf r_int_FilCab = 56 Then
            .Cells(r_int_FilCab, 2) = "C'.- MONTO DE LOS CREDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO CREDITICIO DE LOS CREDITOS INDIRECTOS CON SUSTITUCION DE CONTRAPARTE CREDITICIA - DESPUES DE LA SUSTITUCION 5b/"
         ElseIf r_int_FilCab = 66 Then
            .Cells(r_int_FilCab, 2) = "D.- MONTO DE LOS CREDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO CREDITICIO DE LOS CREDITOS INDIRECTOS QUE CUENTAN CON GARANTIAS PREFERIDAS AUTOLIQUIDABLES 6/"
         ElseIf r_int_FilCab = 76 Then
            .Cells(r_int_FilCab, 2) = "D'.- MONTO DE LOS CREDITOS QUE CUENTAN CON CONVENIOS ELEGIBLES 7/"
         ElseIf r_int_FilCab = 78 Then
            .Cells(r_int_FilCab, 2) = "E.- MONTO DE LOS CREDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO CREDITICIO DE LOS CREDITOS INDIRECTOS QUE CUENTEN CON GARANTIAS PREFERIDAS DE MUY RAPIDA REALIZACION 8/"
         ElseIf r_int_FilCab = 86 Then
            .Cells(r_int_FilCab, 2) = "F.- MONTO DE LOS CREDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO CREDITICIO DE LOS CREDITOS INDIRECTOS QUE CUENTAN CON GARANTIAS PREFERIDAS 9/"
         ElseIf r_int_FilCab = 96 Then
            .Cells(r_int_FilCab, 2) = "G.- MONTO DE LOS CREDITOS HIPOTECARIOS QUE CUENTAN CON COBERTURA DEL FONDO MIVIVIENDA 10/"
         ElseIf r_int_FilCab = 98 Then
            .Cells(r_int_FilCab, 2) = "H.- MONTO DE LOS CREDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO CREDITICIO DE LOS CREDITOS INDIRECTOS QUE NO CUENTAN CON COBERTURA 12/"
         ElseIf r_int_FilCab = 108 Then
            .Cells(r_int_FilCab, 2) = "I.- PROVISIONES CONSTITUIDAS 13/"
         ElseIf r_int_FilCab = 118 Then
            .Cells(r_int_FilCab, 2) = "J.- PROVISIONES REQUERIDAS 14/"
         ElseIf r_int_FilCab = 128 Then
            .Cells(r_int_FilCab, 2) = "K.- SUPERAVIT (DEFICIT)DE PROVISIONES 15/"
         End If
                  
         .Range(.Cells(r_int_FilCab, 3), .Cells(r_int_FilCab, 8)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).VerticalAlignment = xlCenter
         .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Font.Bold = True
         .Range(.Cells(r_int_FilCab, 2), .Cells(r_int_FilCab, 8)).Borders.Color = RGB(0, 0, 0)
            
         .Cells(r_int_FilCab, 3) = "Normal"
         .Cells(r_int_FilCab, 4) = "CPP"
         .Cells(r_int_FilCab, 5) = "Deficiente"
         .Cells(r_int_FilCab, 6) = "Dudoso"
         .Cells(r_int_FilCab, 7) = "Pérdida"
         .Cells(r_int_FilCab, 8) = "Total"
                 
         If r_int_FilCab = 76 Or r_int_FilCab = 96 Then
            r_int_FilCab = r_int_FilCab + 1
         Else
            If r_int_FilCab = 78 Then
               r_int_FilCab = r_int_FilCab + 7
            Else
               r_int_FilCab = r_int_FilCab + 9
            End If
         End If
      Next r_int_FilCab
      
      
      For r_int_FilDet = 17 To 137
         
         If r_int_FilDet <> 79 Then
         
            .Cells(r_int_FilDet + 0, 2) = "Corporativos"
            .Cells(r_int_FilDet + 1, 2) = "Grandes Empresas"
            .Cells(r_int_FilDet + 2, 2) = "Medianas Empresas"
            .Cells(r_int_FilDet + 3, 2) = "Pequeñas Empresas"
            .Cells(r_int_FilDet + 4, 2) = "Microempresas"
            .Cells(r_int_FilDet + 5, 2) = "Consumo revolvente"
            .Cells(r_int_FilDet + 6, 2) = "Consumo no revolvente"
            .Cells(r_int_FilDet + 7, 2) = "Hipotecario para Vivienda"
            .Cells(r_int_FilDet + 8, 2) = "Total"
         Else
            .Cells(r_int_FilDet + 0, 2) = "Corporativos"
            .Cells(r_int_FilDet + 1, 2) = "Grandes Empresas"
            .Cells(r_int_FilDet + 2, 2) = "Medianas Empresas"
            .Cells(r_int_FilDet + 3, 2) = "Pequeñas Empresas"
            .Cells(r_int_FilDet + 4, 2) = "Microempresas"
            .Cells(r_int_FilDet + 5, 2) = "Hipotecario para Vivienda"
            .Cells(r_int_FilDet + 6, 2) = "Total"
         
         End If
         
         For r_int_VarAux = 0 To 8
            .Range(.Cells(r_int_FilDet + r_int_VarAux, 2), .Cells(r_int_FilDet + r_int_VarAux, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(r_int_FilDet + r_int_VarAux, 2), .Cells(r_int_FilDet + r_int_VarAux, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(r_int_FilDet + r_int_VarAux, 2), .Cells(r_int_FilDet + r_int_VarAux, 8)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next r_int_VarAux
                   
         If r_int_FilDet = 67 Or r_int_FilDet = 87 Then
            r_int_FilDet = r_int_FilDet + 11
         Else
            If r_int_FilDet = 79 Then
               r_int_FilDet = r_int_FilDet + 7
            Else
               r_int_FilDet = r_int_FilDet + 9
            End If
         End If
                 
      Next r_int_FilDet
      
      .Range(.Cells(137, 2), .Cells(137, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(24, 2), .Cells(24, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(34, 2), .Cells(34, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(44, 2), .Cells(44, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(54, 2), .Cells(54, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(64, 2), .Cells(64, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(74, 2), .Cells(74, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(84, 2), .Cells(84, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(94, 2), .Cells(94, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      .Range(.Cells(106, 2), .Cells(106, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(116, 2), .Cells(116, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(126, 2), .Cells(126, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(136, 2), .Cells(136, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
      .Range(.Range("B2:H2"), .Range("B14:H14")).Font.Bold = True
      .Range(.Range("B2:H2"), .Range("B14:H14")).HorizontalAlignment = xlHAlignCenter
            
      .Range("B5:H5").Merge
      .Range("B6:H6").Merge
      .Range("B10:H10").Merge
      .Range("B11:H11").Merge
      .Range("B13:H13").Merge
            
      .Range("B4").HorizontalAlignment = xlHAlignLeft
           
      .Range("B16").RowHeight = 40
      .Range("B16").WrapText = True
      
      .Range("B26").RowHeight = 40
      .Range("B26").WrapText = True
      
      .Range("B36").RowHeight = 40
      .Range("B36").WrapText = True
      
      .Range("B46").RowHeight = 40
      .Range("B46").WrapText = True
           
      .Range("B56").RowHeight = 40
      .Range("B56").WrapText = True
           
      .Range("B66").RowHeight = 40
      .Range("B66").WrapText = True
           
      .Range("B76").RowHeight = 40
      .Range("B76").WrapText = True
      
      .Range("B78").RowHeight = 40
      .Range("B78").WrapText = True
           
      .Range("B86").RowHeight = 40
      .Range("B86").WrapText = True
           
      .Range("B96").RowHeight = 40
      .Range("B96").WrapText = True
           
      .Range("B98").RowHeight = 40
      .Range("B98").WrapText = True
      
      .Range("B108").RowHeight = 40
      .Range("B108").WrapText = True
      
      .Range("B118").RowHeight = 40
      .Range("B118").WrapText = True
           
      .Range("B128").RowHeight = 40
      .Range("B128").WrapText = True
      
      .Range(.Cells(97, 1), .Cells(97, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
           
      .Columns("A").ColumnWidth = 2
      .Columns("B").ColumnWidth = 76
      .Columns("C").ColumnWidth = 15
      .Columns("D").ColumnWidth = 15
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 15
      .Columns("G").ColumnWidth = 15
      .Columns("H").ColumnWidth = 15
      .Columns("I").ColumnWidth = 2
      
      .Range("B142:H142").Merge
      .Range("B145:F145").Merge
      .Range("B154:F154").Merge
      
      .Range(.Range("B142:H142"), .Range("B146:H146")).Font.Bold = True
      .Range(.Range("B142:H142"), .Range("B146:H146")).HorizontalAlignment = xlHAlignCenter
      .Range(.Range("B154:H154"), .Range("B155:H155")).HorizontalAlignment = xlHAlignCenter
      .Range(.Range("B154:H154"), .Range("B155:H155")).VerticalAlignment = xlVAlignCenter
      
      .Range("B146:H146").RowHeight = 40
      .Range(.Range("B142:H142"), .Range("B146:H146")).WrapText = True
      
      .Range("B146:F146").VerticalAlignment = xlVAlignCenter
      .Range("B155:E155").VerticalAlignment = xlVAlignCenter
      .Range("B150").VerticalAlignment = xlVAlignCenter
      .Range("B152").VerticalAlignment = xlVAlignCenter
      
      .Range("B150").RowHeight = 30
      .Range("B152").RowHeight = 30
      
      .Range("B150").WrapText = True
      .Range("B152").WrapText = True
      
      .Range("B155:H155").RowHeight = 78
      .Range(.Range("B155:H155"), .Range("B155:H155")).WrapText = True
      .Range(.Range("B154:F154"), .Range("B155:FH155")).Font.Bold = True
      
      .Range(.Range("B145:F145"), .Range("B156:F156")).Borders.Color = RGB(0, 0, 0)
      
      
      .Range(.Cells(1, 1), .Cells(400, 400)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(400, 400)).Font.Size = 10
      
      .Range(.Cells(1, 1), .Cells(35, 9)).NumberFormat = "###,###,##0.00"
      .Range(.Cells(47, 1), .Cells(400, 9)).NumberFormat = "###,###,##0.00"
   
      g_rst_Princi.MoveFirst
        
      r_int_ValArr = 0
         
      For r_int_FilExc = 17 To 137
      
         For r_int_ColExc = 0 To 5
            
            If (r_int_FilExc <> 26) And (r_int_FilExc <> 36) And (r_int_FilExc <> 46) And (r_int_FilExc <> 56) And _
               (r_int_FilExc <> 66) And (r_int_FilExc <> 76) And (r_int_FilExc <> 78) And (r_int_FilExc <> 86) And _
               (r_int_FilExc <> 96) And (r_int_FilExc <> 98) And (r_int_FilExc <> 108) And (r_int_FilExc <> 118) And _
               (r_int_FilExc <> 128) Then
            
               r_obj_Excel.ActiveSheet.Cells(r_int_FilExc, r_int_ColExc + 3) = l_str_Evalua(r_int_ValArr)
            
            r_int_ValArr = r_int_ValArr + 1
            
            End If
            
         Next r_int_ColExc
         
      Next r_int_FilExc
      
      
      r_int_ValArr = 648
         
      For r_int_FilExc = 148 To 156
      
         For r_int_ColExc = 0 To 3
            
            If (r_int_FilExc <> 149) And (r_int_FilExc <> 154) And (r_int_FilExc <> 155) Then
            
               If r_int_ValArr = 668 Then
                  r_int_ColExc = r_int_ColExc + 1
               End If
            
               r_obj_Excel.ActiveSheet.Cells(r_int_FilExc, r_int_ColExc + 3) = l_str_Evalua(r_int_ValArr)
            
               r_int_ValArr = r_int_ValArr + 1
            
            End If
            
         Next r_int_ColExc
         
      Next r_int_FilExc
      
      
      .Range(.Cells(1, 1), .Cells(163, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(163, 99)).Font.Size = 8
            
   End With
         
   Screen.MousePointer = 0
  
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing

End Sub

Private Function ff_ValDat(ByVal p_PerMes As String, ByVal p_PerAno As String) As Integer
   
   ff_ValDat = 0
   
   g_str_Parame = "SELECT COUNT(*) AS TOTAL FROM HIS_CLADEU WHERE "
   g_str_Parame = g_str_Parame & "CLADEU_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "CLADEU_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "CLADEU_NOMREP = 'CTB_REPSBS_05' AND "
   g_str_Parame = g_str_Parame & "CLADEU_TERCRE = '" & modgen_g_str_NombPC & "' "
      
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

Private Sub fs_CalDat()
   
   Dim r_int_Contad           As Integer
   Dim r_int_ConAux           As Integer
   Dim r_int_VarCon           As Integer
   Dim r_int_ConTmp           As Integer
   Dim r_int_ValCad           As Integer
   Dim r_int_CodInd           As Integer
   Dim r_int_CodPer           As Integer
   
   Dim r_dbl_MtoNor(14)       As Double
   Dim r_dbl_MtoCpp(14)       As Double
   Dim r_dbl_MtoDef(14)       As Double
   Dim r_dbl_MtoDud(14)       As Double
   Dim r_dbl_MtoPer(14)       As Double
   
   Dim r_str_CabRep(14)       As String
   Dim r_str_Garlin(14)       As String
   
   Dim c As Integer
   Dim i As Integer
   
   r_int_Contad = 0
   
   Erase l_str_Evalua()
   
   'Leer Tabla de Creditos del mes CRE_HIPCIE
   g_str_Parame = "SELECT HIPCIE_NUMOPE, HIPCIE_TIPGAR, HIPCIE_CLAPRV, HIPCIE_PRVGEN, HIPCIE_PRVESP, HIPCIE_PRVCIC, HIPCIE_SITCRE, HIPCIE_CLAPRV, HIPCIE_SALCAP, HIPCIE_CAPVIG, HIPCIE_SALCON, HIPCIE_TIPCAM, HIPCIE_TIPMON, HIPCIE_CODPRD, HIPCIE_FECDES, HIPCIE_INTDIF FROM CRE_HIPCIE H WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_NUMOPE ASC "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se encontraron Saldos para generar el Reporte.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   g_rst_Princi.MoveFirst
     
   Do While Not g_rst_Princi.EOF
      
      If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
         Call fs_CalHip(426, 204, 216, 486, 540, 414, 474, 150, 528, 582, 42)
      ElseIf g_rst_Princi!HIPCIE_CLAPRV = 1 Then
         Call fs_CalHip(427, 205, 217, 487, 541, 415, 475, 151, 529, 583, 43)
      ElseIf g_rst_Princi!HIPCIE_CLAPRV = 2 Then
         Call fs_CalHip(428, 206, 218, 488, 542, 416, 476, 152, 530, 584, 44)
      ElseIf g_rst_Princi!HIPCIE_CLAPRV = 3 Then
         Call fs_CalHip(429, 207, 219, 489, 543, 417, 477, 153, 531, 585, 45)
      ElseIf g_rst_Princi!HIPCIE_CLAPRV = 4 Then
         Call fs_CalHip(430, 208, 220, 490, 544, 418, 478, 154, 532, 586, 46)
      End If
      
      'If g_rst_Princi!HIPCIE_TIPMON = 1 Then '
      '   r_int_Contad = 47
      '   l_str_Evalua(r_int_Contad) = l_str_Evalua(r_int_Contad) + (g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON)
      'Else
      '   r_int_Contad = 47
      '   l_str_Evalua(r_int_Contad) = l_str_Evalua(r_int_Contad) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
      'End If
      
      'r_int_Contad = 155
      'l_str_Evalua(r_int_Contad) = l_str_Evalua(r_int_Contad) + 1
      
      'r_int_Contad = 587
      'l_str_Evalua(r_int_Contad) = l_str_Evalua(r_int_Contad) + g_rst_Princi!HIPCIE_PRVESP + g_rst_Princi!HIPCIE_PRVGEN
      
      g_rst_Princi.MoveNext
      DoEvents
  
   Loop
       
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   'Leer Tabla de Creditos del mes CRE_COMCIE
   g_str_Parame = "SELECT MAX(COMCIE_NDOCLI) AS NDOCLI, MAX(COMCIE_NUECRE) AS NUECRE, MAX(COMCIE_TIPGAR) AS TIPGAR, MAX(COMCIE_CLAPRV) AS CLAPRV, SUM(COMCIE_PRVGEN) AS PRVGEN, SUM(COMCIE_PRVESP) AS PRVESP, SUM(COMCIE_PRVCIC) AS PRVCIC, MAX(COMCIE_SITCRE) AS SITCRE, MAX(COMCIE_CLAPRV) AS CLAPRV, SUM(COMCIE_SALCAP) AS SALCAP, MAX(COMCIE_TIPCAM) AS TIPCAM, MAX(COMCIE_TIPMON) AS TIPMON FROM CRE_COMCIE H WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
   g_str_Parame = g_str_Parame & "GROUP BY COMCIE_NDOCLI "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No se encontraron Saldos para generar el Reporte.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
      If g_rst_Princi!NUECRE = 6 Then
         Call fs_CalCom(0, 108, 540, 372, 432)
      ElseIf g_rst_Princi!NUECRE = 7 Then
         Call fs_CalCom(6, 114, 546, 378, 438)
      ElseIf g_rst_Princi!NUECRE = 8 Then
         Call fs_CalCom(12, 120, 552, 384, 444)
      ElseIf g_rst_Princi!NUECRE = 9 Then
         Call fs_CalCom(18, 126, 558, 390, 450)
      ElseIf g_rst_Princi!NUECRE = 10 Then
         Call fs_CalCom(24, 132, 564, 396, 456)
      ElseIf g_rst_Princi!NUECRE = 11 Then
         Call fs_CalCom(30, 138, 570, 402, 462)
      ElseIf g_rst_Princi!NUECRE = 12 Then
         Call fs_CalCom(36, 144, 576, 408, 468)
      End If
   
      g_rst_Princi.MoveNext
      DoEvents
  
   Loop
 
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   
   'Totalizando Valores
   c = 0
   
   For i = 48 To 318
   
      l_str_Evalua(c + 5) = l_str_Evalua(c + 0) + l_str_Evalua(c + 1) + l_str_Evalua(c + 2) + l_str_Evalua(c + 3) + l_str_Evalua(c + 4)
      l_str_Evalua(c + 11) = l_str_Evalua(c + 6) + l_str_Evalua(c + 7) + l_str_Evalua(c + 8) + l_str_Evalua(c + 9) + l_str_Evalua(c + 10)
      l_str_Evalua(c + 17) = l_str_Evalua(c + 12) + l_str_Evalua(c + 13) + l_str_Evalua(c + 14) + l_str_Evalua(c + 15) + l_str_Evalua(c + 16)
      l_str_Evalua(c + 23) = l_str_Evalua(c + 18) + l_str_Evalua(c + 19) + l_str_Evalua(c + 20) + l_str_Evalua(c + 21) + l_str_Evalua(c + 22)
      l_str_Evalua(c + 29) = l_str_Evalua(c + 24) + l_str_Evalua(c + 25) + l_str_Evalua(c + 26) + l_str_Evalua(c + 27) + l_str_Evalua(c + 28)
      l_str_Evalua(c + 35) = l_str_Evalua(c + 30) + l_str_Evalua(c + 31) + l_str_Evalua(c + 32) + l_str_Evalua(c + 33) + l_str_Evalua(c + 34)
      l_str_Evalua(c + 41) = l_str_Evalua(c + 36) + l_str_Evalua(c + 37) + l_str_Evalua(c + 38) + l_str_Evalua(c + 39) + l_str_Evalua(c + 40)
      l_str_Evalua(c + 47) = l_str_Evalua(c + 42) + l_str_Evalua(c + 43) + l_str_Evalua(c + 44) + l_str_Evalua(c + 45) + l_str_Evalua(c + 46)
  
      l_str_Evalua(i + 0) = l_str_Evalua(c + 0) + l_str_Evalua(c + 6) + l_str_Evalua(c + 12) + l_str_Evalua(c + 18) + l_str_Evalua(c + 24) + l_str_Evalua(c + 30) + l_str_Evalua(c + 36) + l_str_Evalua(c + 42)
      l_str_Evalua(i + 1) = l_str_Evalua(c + 1) + l_str_Evalua(c + 7) + l_str_Evalua(c + 13) + l_str_Evalua(c + 19) + l_str_Evalua(c + 25) + l_str_Evalua(c + 31) + l_str_Evalua(c + 37) + l_str_Evalua(c + 43)
      l_str_Evalua(i + 2) = l_str_Evalua(c + 2) + l_str_Evalua(c + 8) + l_str_Evalua(c + 14) + l_str_Evalua(c + 20) + l_str_Evalua(c + 26) + l_str_Evalua(c + 32) + l_str_Evalua(c + 38) + l_str_Evalua(c + 44)
      l_str_Evalua(i + 3) = l_str_Evalua(c + 3) + l_str_Evalua(c + 9) + l_str_Evalua(c + 15) + l_str_Evalua(c + 21) + l_str_Evalua(c + 27) + l_str_Evalua(c + 33) + l_str_Evalua(c + 39) + l_str_Evalua(c + 45)
      l_str_Evalua(i + 4) = l_str_Evalua(c + 4) + l_str_Evalua(c + 10) + l_str_Evalua(c + 16) + l_str_Evalua(c + 22) + l_str_Evalua(c + 28) + l_str_Evalua(c + 34) + l_str_Evalua(c + 40) + l_str_Evalua(c + 46)
      l_str_Evalua(i + 5) = l_str_Evalua(c + 5) + l_str_Evalua(c + 11) + l_str_Evalua(c + 17) + l_str_Evalua(c + 23) + l_str_Evalua(c + 29) + l_str_Evalua(c + 35) + l_str_Evalua(c + 41) + l_str_Evalua(c + 47)
      
      c = c + 54
      i = i + 53
      
   Next i
   
   c = 330
   
   l_str_Evalua(335) = l_str_Evalua(c + 0) + l_str_Evalua(c + 1) + l_str_Evalua(c + 2) + l_str_Evalua(c + 3) + l_str_Evalua(c + 4)
   l_str_Evalua(341) = l_str_Evalua(c + 6) + l_str_Evalua(c + 7) + l_str_Evalua(c + 8) + l_str_Evalua(c + 9) + l_str_Evalua(c + 10)
   l_str_Evalua(347) = l_str_Evalua(c + 12) + l_str_Evalua(c + 13) + l_str_Evalua(c + 14) + l_str_Evalua(c + 15) + l_str_Evalua(c + 16)
   l_str_Evalua(353) = l_str_Evalua(c + 18) + l_str_Evalua(c + 19) + l_str_Evalua(c + 20) + l_str_Evalua(c + 21) + l_str_Evalua(c + 22)
   l_str_Evalua(359) = l_str_Evalua(c + 24) + l_str_Evalua(c + 25) + l_str_Evalua(c + 26) + l_str_Evalua(c + 27) + l_str_Evalua(c + 28)
   l_str_Evalua(365) = l_str_Evalua(c + 30) + l_str_Evalua(c + 31) + l_str_Evalua(c + 32) + l_str_Evalua(c + 33) + l_str_Evalua(c + 34)
   l_str_Evalua(371) = l_str_Evalua(c + 36) + l_str_Evalua(c + 37) + l_str_Evalua(c + 38) + l_str_Evalua(c + 39) + l_str_Evalua(c + 40)
      
   l_str_Evalua(366) = l_str_Evalua(c + 0) + l_str_Evalua(c + 6) + l_str_Evalua(c + 12) + l_str_Evalua(c + 18) + l_str_Evalua(c + 24) + l_str_Evalua(c + 30) + l_str_Evalua(c + 36) + l_str_Evalua(c + 42)
   l_str_Evalua(367) = l_str_Evalua(c + 1) + l_str_Evalua(c + 7) + l_str_Evalua(c + 13) + l_str_Evalua(c + 19) + l_str_Evalua(c + 25) + l_str_Evalua(c + 31) + l_str_Evalua(c + 37) + l_str_Evalua(c + 43)
   l_str_Evalua(368) = l_str_Evalua(c + 2) + l_str_Evalua(c + 8) + l_str_Evalua(c + 14) + l_str_Evalua(c + 20) + l_str_Evalua(c + 26) + l_str_Evalua(c + 32) + l_str_Evalua(c + 38) + l_str_Evalua(c + 44)
   l_str_Evalua(369) = l_str_Evalua(c + 3) + l_str_Evalua(c + 9) + l_str_Evalua(c + 15) + l_str_Evalua(c + 21) + l_str_Evalua(c + 27) + l_str_Evalua(c + 33) + l_str_Evalua(c + 39) + l_str_Evalua(c + 45)
   l_str_Evalua(370) = l_str_Evalua(c + 4) + l_str_Evalua(c + 10) + l_str_Evalua(c + 16) + l_str_Evalua(c + 22) + l_str_Evalua(c + 28) + l_str_Evalua(c + 34) + l_str_Evalua(c + 40) + l_str_Evalua(c + 46)
   l_str_Evalua(371) = l_str_Evalua(c + 5) + l_str_Evalua(c + 11) + l_str_Evalua(c + 17) + l_str_Evalua(c + 23) + l_str_Evalua(c + 29) + l_str_Evalua(c + 35) + l_str_Evalua(c + 41) + l_str_Evalua(c + 47)
   
   c = 372
   
   l_str_Evalua(377) = l_str_Evalua(c + 0) + l_str_Evalua(c + 1) + l_str_Evalua(c + 2) + l_str_Evalua(c + 3) + l_str_Evalua(c + 4)
   l_str_Evalua(383) = l_str_Evalua(c + 6) + l_str_Evalua(c + 7) + l_str_Evalua(c + 8) + l_str_Evalua(c + 9) + l_str_Evalua(c + 10)
   l_str_Evalua(389) = l_str_Evalua(c + 12) + l_str_Evalua(c + 13) + l_str_Evalua(c + 14) + l_str_Evalua(c + 15) + l_str_Evalua(c + 16)
   l_str_Evalua(395) = l_str_Evalua(c + 18) + l_str_Evalua(c + 19) + l_str_Evalua(c + 20) + l_str_Evalua(c + 21) + l_str_Evalua(c + 22)
   l_str_Evalua(401) = l_str_Evalua(c + 24) + l_str_Evalua(c + 25) + l_str_Evalua(c + 26) + l_str_Evalua(c + 27) + l_str_Evalua(c + 28)
   l_str_Evalua(407) = l_str_Evalua(c + 30) + l_str_Evalua(c + 31) + l_str_Evalua(c + 32) + l_str_Evalua(c + 33) + l_str_Evalua(c + 34)
   l_str_Evalua(413) = l_str_Evalua(c + 36) + l_str_Evalua(c + 37) + l_str_Evalua(c + 38) + l_str_Evalua(c + 39) + l_str_Evalua(c + 40)
   l_str_Evalua(419) = l_str_Evalua(c + 42) + l_str_Evalua(c + 43) + l_str_Evalua(c + 44) + l_str_Evalua(c + 45) + l_str_Evalua(c + 46)
   
   l_str_Evalua(420) = l_str_Evalua(c + 0) + l_str_Evalua(c + 6) + l_str_Evalua(c + 12) + l_str_Evalua(c + 18) + l_str_Evalua(c + 24) + l_str_Evalua(c + 30) + l_str_Evalua(c + 36) + l_str_Evalua(c + 42)
   l_str_Evalua(421) = l_str_Evalua(c + 1) + l_str_Evalua(c + 7) + l_str_Evalua(c + 13) + l_str_Evalua(c + 19) + l_str_Evalua(c + 25) + l_str_Evalua(c + 31) + l_str_Evalua(c + 37) + l_str_Evalua(c + 43)
   l_str_Evalua(422) = l_str_Evalua(c + 2) + l_str_Evalua(c + 8) + l_str_Evalua(c + 14) + l_str_Evalua(c + 20) + l_str_Evalua(c + 26) + l_str_Evalua(c + 32) + l_str_Evalua(c + 38) + l_str_Evalua(c + 44)
   l_str_Evalua(423) = l_str_Evalua(c + 3) + l_str_Evalua(c + 9) + l_str_Evalua(c + 15) + l_str_Evalua(c + 21) + l_str_Evalua(c + 27) + l_str_Evalua(c + 33) + l_str_Evalua(c + 39) + l_str_Evalua(c + 45)
   l_str_Evalua(424) = l_str_Evalua(c + 4) + l_str_Evalua(c + 10) + l_str_Evalua(c + 16) + l_str_Evalua(c + 22) + l_str_Evalua(c + 28) + l_str_Evalua(c + 34) + l_str_Evalua(c + 40) + l_str_Evalua(c + 46)
   l_str_Evalua(425) = l_str_Evalua(c + 5) + l_str_Evalua(c + 11) + l_str_Evalua(c + 17) + l_str_Evalua(c + 23) + l_str_Evalua(c + 29) + l_str_Evalua(c + 35) + l_str_Evalua(c + 41) + l_str_Evalua(c + 47)
   
   l_str_Evalua(431) = l_str_Evalua(c + 54) + l_str_Evalua(c + 55) + l_str_Evalua(c + 56) + l_str_Evalua(c + 57) + l_str_Evalua(c + 58)
   
   
   
   
   If l_int_PrvPro = 1 Then
   
      g_str_Parame = "SELECT * FROM CTB_MNTIND WHERE MNTIND_CODIGO = '001'"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros para generar el Reporte.", vbInformation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
            
      Do While Not g_rst_Princi.EOF
         
         r_int_CodInd = Trim(g_rst_Princi!MNTIND_CODIND)
         r_int_CodPer = IIf(IsNull(Trim(g_rst_Princi!MNTIND_CODPER)) = True, 0, Trim(g_rst_Princi!MNTIND_CODPER))
            
         g_rst_Princi.MoveNext
         DoEvents
     
      Loop
    
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      If r_int_CodInd = 1 And r_int_CodPer = 0 Then
         l_str_Evalua(540) = l_str_Evalua(540) + Format((l_str_Evalua(54) - l_str_Evalua(270) - l_str_Evalua(162) + l_str_Evalua(216)) * 0, "###,###,##0.00")
   
         l_str_Evalua(546) = l_str_Evalua(546) + Format((l_str_Evalua(60) - l_str_Evalua(276) - l_str_Evalua(168) + l_str_Evalua(222)) * 0, "###,###,##0.00")
   
         l_str_Evalua(552) = l_str_Evalua(552) + Format((l_str_Evalua(66) - l_str_Evalua(282) - l_str_Evalua(174)) * 0, "###,###,##0.00")
         l_str_Evalua(558) = l_str_Evalua(558) + Format((l_str_Evalua(72) - l_str_Evalua(288) - l_str_Evalua(180)) * 0, "###,###,##0.00")
         l_str_Evalua(564) = l_str_Evalua(564) + Format((l_str_Evalua(78) - l_str_Evalua(294) - l_str_Evalua(186)) * 0, "###,###,##0.00")
         l_str_Evalua(570) = l_str_Evalua(570) + Format((l_str_Evalua(84) - l_str_Evalua(300) - l_str_Evalua(192)) * 0, "###,###,##0.00")
         l_str_Evalua(576) = l_str_Evalua(576) + Format((l_str_Evalua(90) - l_str_Evalua(360) - l_str_Evalua(324) - l_str_Evalua(198)) * 0, "###,###,##0.00")
   
         l_str_Evalua(582) = l_str_Evalua(582) + Format((l_str_Evalua(96) - l_str_Evalua(312) - l_str_Evalua(426) - l_str_Evalua(204)) * 0, "###,###,##0.00")
   
   
      ElseIf r_int_CodInd = 1 And r_int_CodPer = 1 Then
         l_str_Evalua(540) = l_str_Evalua(540) + Format((l_str_Evalua(54) - l_str_Evalua(270) - l_str_Evalua(162) + l_str_Evalua(216)) * 0.0015, "###,###,##0.00")
   
         l_str_Evalua(546) = l_str_Evalua(546) + Format((l_str_Evalua(60) - l_str_Evalua(276) - l_str_Evalua(168) + l_str_Evalua(222)) * 0.0015, "###,###,##0.00")
   
         l_str_Evalua(552) = l_str_Evalua(552) + Format((l_str_Evalua(66) - l_str_Evalua(282) - l_str_Evalua(174)) * 0.001, "###,###,##0.00")
         l_str_Evalua(558) = l_str_Evalua(558) + Format((l_str_Evalua(72) - l_str_Evalua(288) - l_str_Evalua(180)) * 0.002, "###,###,##0.00")
         l_str_Evalua(564) = l_str_Evalua(564) + Format((l_str_Evalua(78) - l_str_Evalua(294) - l_str_Evalua(186)) * 0.002, "###,###,##0.00")
         l_str_Evalua(570) = l_str_Evalua(570) + Format((l_str_Evalua(84) - l_str_Evalua(300) - l_str_Evalua(192)) * 0.005, "###,###,##0.00")
         l_str_Evalua(576) = l_str_Evalua(576) + Format((l_str_Evalua(90) - l_str_Evalua(360) - l_str_Evalua(324) - l_str_Evalua(198)) * 0.004, "###,###,##0.00")
   
         l_str_Evalua(582) = l_str_Evalua(582) + Format((l_str_Evalua(96) - l_str_Evalua(312) - l_str_Evalua(426) - l_str_Evalua(204)) * 0.0015, "###,###,##0.00")
   
   
      ElseIf r_int_CodInd = 1 And r_int_CodPer = 2 Then
         l_str_Evalua(540) = l_str_Evalua(540) + Format((l_str_Evalua(54) - l_str_Evalua(270) - l_str_Evalua(162) + l_str_Evalua(216)) * 0.003, "###,###,##0.00")
   
         l_str_Evalua(546) = l_str_Evalua(546) + Format((l_str_Evalua(60) - l_str_Evalua(276) - l_str_Evalua(168) + l_str_Evalua(222)) * 0.003, "###,###,##0.00")
   
         l_str_Evalua(552) = l_str_Evalua(552) + Format((l_str_Evalua(66) - l_str_Evalua(282) - l_str_Evalua(174)) * 0.002, "###,###,##0.00")
         l_str_Evalua(558) = l_str_Evalua(558) + Format((l_str_Evalua(72) - l_str_Evalua(288) - l_str_Evalua(180)) * 0.004, "###,###,##0.00")
         l_str_Evalua(564) = l_str_Evalua(564) + Format((l_str_Evalua(78) - l_str_Evalua(294) - l_str_Evalua(186)) * 0.004, "###,###,##0.00")
         l_str_Evalua(570) = l_str_Evalua(570) + Format((l_str_Evalua(84) - l_str_Evalua(300) - l_str_Evalua(192)) * 0.01, "###,###,##0.00")
         l_str_Evalua(576) = l_str_Evalua(576) + Format((l_str_Evalua(90) - l_str_Evalua(360) - l_str_Evalua(324) - l_str_Evalua(198)) * 0.007, "###,###,##0.00")
   
         l_str_Evalua(582) = l_str_Evalua(582) + Format((l_str_Evalua(96) - l_str_Evalua(312) - l_str_Evalua(426) - l_str_Evalua(204)) * 0.003, "###,###,##0.00")
   
   
      ElseIf r_int_CodInd = 0 Then
         l_str_Evalua(540) = l_str_Evalua(540) + Format((l_str_Evalua(54) - l_str_Evalua(270) - l_str_Evalua(162) + l_str_Evalua(216)) * 0.004, "###,###,##0.00")
   
         l_str_Evalua(546) = l_str_Evalua(546) + Format((l_str_Evalua(60) - l_str_Evalua(276) - l_str_Evalua(168) + l_str_Evalua(222)) * 0.0045, "###,###,##0.00")
   
         l_str_Evalua(552) = l_str_Evalua(552) + Format((l_str_Evalua(66) - l_str_Evalua(282) - l_str_Evalua(174)) * 0.003, "###,###,##0.00")
         l_str_Evalua(558) = l_str_Evalua(558) + Format((l_str_Evalua(72) - l_str_Evalua(288) - l_str_Evalua(180)) * 0.005, "###,###,##0.00")
         l_str_Evalua(564) = l_str_Evalua(564) + Format((l_str_Evalua(78) - l_str_Evalua(294) - l_str_Evalua(186)) * 0.005, "###,###,##0.00")
         l_str_Evalua(570) = l_str_Evalua(570) + Format((l_str_Evalua(84) - l_str_Evalua(300) - l_str_Evalua(192)) * 0.015, "###,###,##0.00")
         l_str_Evalua(576) = l_str_Evalua(576) + Format((l_str_Evalua(90) - l_str_Evalua(360) - l_str_Evalua(324) - l_str_Evalua(198)) * 0.01, "###,###,##0.00")
   
         l_str_Evalua(582) = l_str_Evalua(582) + Format((l_str_Evalua(96) - l_str_Evalua(312) - l_str_Evalua(426) - l_str_Evalua(204)) * 0.004, "###,###,##0.00")
      
      End If
   
   End If
   
   
   
   
   For i = 594 To 641
   
      l_str_Evalua(i) = l_str_Evalua(i - 108) - l_str_Evalua(i - 54)
      
   Next i
      
   c = 432
   
   For i = 480 To 642
   
      l_str_Evalua(c + 5) = l_str_Evalua(c + 0) + l_str_Evalua(c + 1) + l_str_Evalua(c + 2) + l_str_Evalua(c + 3) + l_str_Evalua(c + 4)
      l_str_Evalua(c + 11) = l_str_Evalua(c + 6) + l_str_Evalua(c + 7) + l_str_Evalua(c + 8) + l_str_Evalua(c + 9) + l_str_Evalua(c + 10)
      l_str_Evalua(c + 17) = l_str_Evalua(c + 12) + l_str_Evalua(c + 13) + l_str_Evalua(c + 14) + l_str_Evalua(c + 15) + l_str_Evalua(c + 16)
      l_str_Evalua(c + 23) = l_str_Evalua(c + 18) + l_str_Evalua(c + 19) + l_str_Evalua(c + 20) + l_str_Evalua(c + 21) + l_str_Evalua(c + 22)
      l_str_Evalua(c + 29) = l_str_Evalua(c + 24) + l_str_Evalua(c + 25) + l_str_Evalua(c + 26) + l_str_Evalua(c + 27) + l_str_Evalua(c + 28)
      l_str_Evalua(c + 35) = l_str_Evalua(c + 30) + l_str_Evalua(c + 31) + l_str_Evalua(c + 32) + l_str_Evalua(c + 33) + l_str_Evalua(c + 34)
      l_str_Evalua(c + 41) = l_str_Evalua(c + 36) + l_str_Evalua(c + 37) + l_str_Evalua(c + 38) + l_str_Evalua(c + 39) + l_str_Evalua(c + 40)
      l_str_Evalua(c + 47) = l_str_Evalua(c + 42) + l_str_Evalua(c + 43) + l_str_Evalua(c + 44) + l_str_Evalua(c + 45) + l_str_Evalua(c + 46)
  
      l_str_Evalua(i + 0) = l_str_Evalua(c + 0) + l_str_Evalua(c + 6) + l_str_Evalua(c + 12) + l_str_Evalua(c + 18) + l_str_Evalua(c + 24) + l_str_Evalua(c + 30) + l_str_Evalua(c + 36) + l_str_Evalua(c + 42)
      l_str_Evalua(i + 1) = l_str_Evalua(c + 1) + l_str_Evalua(c + 7) + l_str_Evalua(c + 13) + l_str_Evalua(c + 19) + l_str_Evalua(c + 25) + l_str_Evalua(c + 31) + l_str_Evalua(c + 37) + l_str_Evalua(c + 43)
      l_str_Evalua(i + 2) = l_str_Evalua(c + 2) + l_str_Evalua(c + 8) + l_str_Evalua(c + 14) + l_str_Evalua(c + 20) + l_str_Evalua(c + 26) + l_str_Evalua(c + 32) + l_str_Evalua(c + 38) + l_str_Evalua(c + 44)
      l_str_Evalua(i + 3) = l_str_Evalua(c + 3) + l_str_Evalua(c + 9) + l_str_Evalua(c + 15) + l_str_Evalua(c + 21) + l_str_Evalua(c + 27) + l_str_Evalua(c + 33) + l_str_Evalua(c + 39) + l_str_Evalua(c + 45)
      l_str_Evalua(i + 4) = l_str_Evalua(c + 4) + l_str_Evalua(c + 10) + l_str_Evalua(c + 16) + l_str_Evalua(c + 22) + l_str_Evalua(c + 28) + l_str_Evalua(c + 34) + l_str_Evalua(c + 40) + l_str_Evalua(c + 46)
      l_str_Evalua(i + 5) = l_str_Evalua(c + 5) + l_str_Evalua(c + 11) + l_str_Evalua(c + 17) + l_str_Evalua(c + 23) + l_str_Evalua(c + 29) + l_str_Evalua(c + 35) + l_str_Evalua(c + 41) + l_str_Evalua(c + 47)
      
      c = c + 54
      i = i + 53
      
   Next i
  
   
   'For c = 48 To 318
   '   l_str_Evalua(c + 5) = l_str_Evalua(c + 0) + l_str_Evalua(c + 1) + l_str_Evalua(c + 2) + l_str_Evalua(c + 3) + l_str_Evalua(c + 4)
   '   c = c + 53
   'Next c
     
   'c = 366
   'l_str_Evalua(c + 5) = l_str_Evalua(c + 0) + l_str_Evalua(c + 1) + l_str_Evalua(c + 2) + l_str_Evalua(c + 3) + l_str_Evalua(c + 4)
      
   'c = 420
   'l_str_Evalua(c + 5) = l_str_Evalua(c + 0) + l_str_Evalua(c + 1) + l_str_Evalua(c + 2) + l_str_Evalua(c + 3) + l_str_Evalua(c + 4)
   
   'For c = 480 To 642
   '   l_str_Evalua(c + 5) = l_str_Evalua(c + 0) + l_str_Evalua(c + 1) + l_str_Evalua(c + 2) + l_str_Evalua(c + 3) + l_str_Evalua(c + 4)
   '   c = c + 53
   'Next c
   
   l_str_Evalua(648) = l_str_Evalua(53)
   l_str_Evalua(649) = l_str_Evalua(53)
   l_str_Evalua(650) = l_str_Evalua(534)
   l_str_Evalua(651) = l_str_Evalua(535) + l_str_Evalua(536) + l_str_Evalua(537) + l_str_Evalua(538)
   
   l_str_Evalua(664) = l_str_Evalua(648) + l_str_Evalua(652) + l_str_Evalua(656) + l_str_Evalua(660)
   l_str_Evalua(665) = l_str_Evalua(649) + l_str_Evalua(653) + l_str_Evalua(657) + l_str_Evalua(661)
   l_str_Evalua(666) = l_str_Evalua(650) + l_str_Evalua(654) + l_str_Evalua(658) + l_str_Evalua(662)
   l_str_Evalua(667) = l_str_Evalua(651) + l_str_Evalua(655) + l_str_Evalua(659) + l_str_Evalua(663)
   
   l_str_Evalua(668) = l_str_Evalua(649)
   l_str_Evalua(669) = l_str_Evalua(650)
   l_str_Evalua(670) = l_str_Evalua(651)
         
   r_int_ConAux = 0
   r_int_ConTmp = 0
   r_int_ValCad = 0
   r_int_ConTmp = 0
   r_int_ConAux = 0
    
   'Insertando a la Tabla
   For r_int_ConTmp = 0 To 107
         
      g_str_Parame = "USP_HIS_CLADEU ("
      g_str_Parame = g_str_Parame & "'CTB_REPSBS_05', "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & CInt(l_str_PerMes) & ", "
      g_str_Parame = g_str_Parame & CInt(l_str_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(Format(r_int_ConTmp, "0000")) & "', "
      g_str_Parame = g_str_Parame & 13 & ", "
               
      For r_int_ConAux = 0 To 5
      
         If r_int_ValCad = (5 * (r_int_ConTmp + 1)) + r_int_ConTmp Then
            g_str_Parame = g_str_Parame & ", " & Format(l_str_Evalua(r_int_ValCad), "###########0.00")
         Else
            g_str_Parame = g_str_Parame & ", " & Format(l_str_Evalua(r_int_ValCad), "###########0.00") & ", "
         End If
                  
         r_int_ValCad = r_int_ValCad + 1
                        
      Next r_int_ConAux
      
      g_str_Parame = g_str_Parame & ")"
               
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_HIS_CLADEU.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
      
   Next r_int_ConTmp
      
End Sub

Sub fs_CalHip(ByVal r_int_Pos001 As Integer, ByVal r_int_Pos002 As Integer, ByVal r_int_Pos003 As Integer, ByVal r_int_Pos004 As Integer, ByVal r_int_Pos005 As Integer, ByVal r_int_Pos006 As Integer, ByVal r_int_Pos007 As Integer, ByVal r_int_Pos008 As Integer, ByVal r_int_Pos009 As Integer, ByVal r_int_Pos010 As Integer, ByVal r_int_Pos011 As Integer)
   
   If g_rst_Princi!HIPCIE_CODPRD <> "002" And g_rst_Princi!HIPCIE_CODPRD <> "005" And g_rst_Princi!HIPCIE_CODPRD <> "006" And g_rst_Princi!HIPCIE_CODPRD <> "008" And g_rst_Princi!HIPCIE_CODPRD <> "011" Then
      If g_rst_Princi!HIPCIE_FECDES <= 20100630 Then
         If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            '426
            l_str_Evalua(r_int_Pos001) = l_str_Evalua(r_int_Pos001) + ((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) / 3)
         Else
            '426
            l_str_Evalua(r_int_Pos001) = l_str_Evalua(r_int_Pos001) + Format(((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM) / 3, "###,###,##0.00")
         End If
      Else
         If g_rst_Princi!HIPCIE_CODPRD = "007" Or g_rst_Princi!HIPCIE_CODPRD = "010" Then
            If g_rst_Princi!HIPCIE_TIPMON = 1 Then
               '204
               l_str_Evalua(r_int_Pos002) = l_str_Evalua(r_int_Pos002) + (g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF)
               '216
               l_str_Evalua(r_int_Pos003) = l_str_Evalua(r_int_Pos003) + (g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF)
            Else
               '204
               l_str_Evalua(r_int_Pos002) = l_str_Evalua(r_int_Pos002) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               '216
               l_str_Evalua(r_int_Pos003) = l_str_Evalua(r_int_Pos003) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            End If
         End If
      
      End If
   End If
   
   If (g_rst_Princi!HIPCIE_TIPGAR = 1 Or g_rst_Princi!HIPCIE_TIPGAR = 2) Then
      If (g_rst_Princi!HIPCIE_CODPRD = "001" Or g_rst_Princi!HIPCIE_CODPRD = "003" Or g_rst_Princi!HIPCIE_CODPRD = "004" Or _
         g_rst_Princi!HIPCIE_CODPRD = "007" Or g_rst_Princi!HIPCIE_CODPRD = "009" Or g_rst_Princi!HIPCIE_CODPRD = "010") Then
         If CDate(gf_FormatoFecha(g_rst_Princi!HIPCIE_FECDES)) <= CDate(gf_FormatoFecha(20100630)) Then
            If g_rst_Princi!HIPCIE_TIPMON = 1 Then
               '414
               l_str_Evalua(r_int_Pos006) = l_str_Evalua(r_int_Pos006) + Format((((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) - g_rst_Princi!HIPCIE_INTDIF) * 2 / 3), "###,###,##0.00")
            Else
               '414
               l_str_Evalua(r_int_Pos006) = l_str_Evalua(r_int_Pos006) + Format(((((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) - g_rst_Princi!HIPCIE_INTDIF) * 2 / 3) * g_rst_Princi!HIPCIE_TIPCAM), "###,###,##0.00")
            End If
         End If

      Else

         If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            '414
            l_str_Evalua(r_int_Pos006) = l_str_Evalua(r_int_Pos006) + ((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) - g_rst_Princi!HIPCIE_INTDIF)
         Else
            '414
            l_str_Evalua(r_int_Pos006) = l_str_Evalua(r_int_Pos006) + Format(((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
         End If

      End If
   Else
      If (g_rst_Princi!HIPCIE_CODPRD = "001" Or g_rst_Princi!HIPCIE_CODPRD = "003" Or g_rst_Princi!HIPCIE_CODPRD = "004" Or _
         g_rst_Princi!HIPCIE_CODPRD = "007" Or g_rst_Princi!HIPCIE_CODPRD = "009" Or g_rst_Princi!HIPCIE_CODPRD = "010") Then
         If CDate(gf_FormatoFecha(g_rst_Princi!HIPCIE_FECDES)) <= CDate(gf_FormatoFecha(20100630)) Then
            If g_rst_Princi!HIPCIE_TIPMON = 1 Then
               '474
               l_str_Evalua(r_int_Pos007) = l_str_Evalua(r_int_Pos007) + Format((((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) - g_rst_Princi!HIPCIE_INTDIF) * 2 / 3), "###,###,##0.00")
            Else
               '474
               l_str_Evalua(r_int_Pos007) = l_str_Evalua(r_int_Pos007) + Format(((((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) - g_rst_Princi!HIPCIE_INTDIF) * 2 / 3) * g_rst_Princi!HIPCIE_TIPCAM), "###,###,##0.00")
            End If
         End If
      Else
         If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            '474
            l_str_Evalua(r_int_Pos007) = l_str_Evalua(r_int_Pos007) + ((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) - g_rst_Princi!HIPCIE_INTDIF)
         Else
            '474
            l_str_Evalua(r_int_Pos007) = l_str_Evalua(r_int_Pos007) + Format(((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
         End If
      End If
   End If
   
   If CDate(gf_FormatoFecha(g_rst_Princi!HIPCIE_FECDES)) >= CDate(gf_FormatoFecha(20100701)) Then

      If (g_rst_Princi!HIPCIE_CODPRD = "001" Or g_rst_Princi!HIPCIE_CODPRD = "003" Or g_rst_Princi!HIPCIE_CODPRD = "004" Or _
         g_rst_Princi!HIPCIE_CODPRD = "007" Or g_rst_Princi!HIPCIE_CODPRD = "009" Or g_rst_Princi!HIPCIE_CODPRD = "010") Then

         If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            '486
            l_str_Evalua(r_int_Pos004) = l_str_Evalua(r_int_Pos004) + (g_rst_Princi!HIPCIE_PRVGEN)
            '540
            l_str_Evalua(r_int_Pos005) = l_str_Evalua(r_int_Pos005) + (g_rst_Princi!HIPCIE_PRVGEN)
         Else
            '486
            l_str_Evalua(r_int_Pos004) = l_str_Evalua(r_int_Pos004) + Format((g_rst_Princi!HIPCIE_PRVGEN) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            '540
            l_str_Evalua(r_int_Pos005) = l_str_Evalua(r_int_Pos005) + Format((g_rst_Princi!HIPCIE_PRVGEN) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
         End If

      End If

   End If
               
   '150
   l_str_Evalua(r_int_Pos008) = l_str_Evalua(r_int_Pos008) + 1
   
   If (g_rst_Princi!HIPCIE_CODPRD = "001" Or g_rst_Princi!HIPCIE_CODPRD = "003" Or g_rst_Princi!HIPCIE_CODPRD = "004" Or _
      g_rst_Princi!HIPCIE_CODPRD = "007" Or g_rst_Princi!HIPCIE_CODPRD = "009" Or g_rst_Princi!HIPCIE_CODPRD = "010") Then
      
      If CDate(gf_FormatoFecha(g_rst_Princi!HIPCIE_FECDES)) <= CDate(gf_FormatoFecha(20100630)) Then
         If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
               '528
               l_str_Evalua(r_int_Pos009) = l_str_Evalua(r_int_Pos009) + g_rst_Princi!HIPCIE_PRVGEN
               '582
               l_str_Evalua(r_int_Pos010) = l_str_Evalua(r_int_Pos010) + g_rst_Princi!HIPCIE_PRVGEN
            Else
               '528
               l_str_Evalua(r_int_Pos009) = l_str_Evalua(r_int_Pos009) + g_rst_Princi!HIPCIE_PRVESP
               '582
               l_str_Evalua(r_int_Pos010) = l_str_Evalua(r_int_Pos010) + g_rst_Princi!HIPCIE_PRVESP
            
            End If
         Else
            If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
               '528
               l_str_Evalua(r_int_Pos009) = l_str_Evalua(r_int_Pos009) + Format(g_rst_Princi!HIPCIE_PRVGEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               '582
               l_str_Evalua(r_int_Pos010) = l_str_Evalua(r_int_Pos010) + Format(g_rst_Princi!HIPCIE_PRVGEN * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            Else
               '528
               l_str_Evalua(r_int_Pos009) = l_str_Evalua(r_int_Pos009) + Format(g_rst_Princi!HIPCIE_PRVESP * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
               '582
               l_str_Evalua(r_int_Pos010) = l_str_Evalua(r_int_Pos010) + Format(g_rst_Princi!HIPCIE_PRVESP * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            
            End If
         End If
      End If
   Else
      If g_rst_Princi!HIPCIE_TIPMON = 1 Then
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            '528
            l_str_Evalua(r_int_Pos009) = l_str_Evalua(r_int_Pos009) + g_rst_Princi!HIPCIE_PRVGEN + IIf(l_int_PrvPro = 1, g_rst_Princi!HIPCIE_PRVCIC, 0)
            '582
            l_str_Evalua(r_int_Pos010) = l_str_Evalua(r_int_Pos010) + g_rst_Princi!HIPCIE_PRVGEN + IIf(l_int_PrvPro = 1, g_rst_Princi!HIPCIE_PRVCIC, 0)
         Else
            '528
            l_str_Evalua(r_int_Pos009) = l_str_Evalua(r_int_Pos009) + g_rst_Princi!HIPCIE_PRVESP + IIf(l_int_PrvPro = 1, g_rst_Princi!HIPCIE_PRVCIC, 0)
            '582
            l_str_Evalua(r_int_Pos010) = l_str_Evalua(r_int_Pos010) + g_rst_Princi!HIPCIE_PRVESP + IIf(l_int_PrvPro = 1, g_rst_Princi!HIPCIE_PRVCIC, 0)
         
         End If
      Else
         If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
            '528
            l_str_Evalua(r_int_Pos009) = l_str_Evalua(r_int_Pos009) + Format((g_rst_Princi!HIPCIE_PRVGEN + IIf(l_int_PrvPro = 1, g_rst_Princi!HIPCIE_PRVCIC, 0)) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            '582
            l_str_Evalua(r_int_Pos010) = l_str_Evalua(r_int_Pos010) + Format((g_rst_Princi!HIPCIE_PRVGEN + IIf(l_int_PrvPro = 1, g_rst_Princi!HIPCIE_PRVCIC, 0)) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
         Else
            '528
            l_str_Evalua(r_int_Pos009) = l_str_Evalua(r_int_Pos009) + Format((g_rst_Princi!HIPCIE_PRVESP + IIf(l_int_PrvPro = 1, g_rst_Princi!HIPCIE_PRVCIC, 0)) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
            '582
            l_str_Evalua(r_int_Pos010) = l_str_Evalua(r_int_Pos010) + Format((g_rst_Princi!HIPCIE_PRVESP + IIf(l_int_PrvPro = 1, g_rst_Princi!HIPCIE_PRVCIC, 0)) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
         
         End If
      End If
   End If
   
   '42
   If g_rst_Princi!HIPCIE_TIPMON = 1 Then
      l_str_Evalua(r_int_Pos011) = l_str_Evalua(r_int_Pos011) + (g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF)
      l_str_Evalua(r_int_Pos011 + 54) = l_str_Evalua(r_int_Pos011 + 54) + (g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF)
   Else
      l_str_Evalua(r_int_Pos011) = l_str_Evalua(r_int_Pos011) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
      l_str_Evalua(r_int_Pos011 + 54) = l_str_Evalua(r_int_Pos011 + 54) + Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON - g_rst_Princi!HIPCIE_INTDIF) * g_rst_Princi!HIPCIE_TIPCAM, "###,###,##0.00")
   End If
      
   
End Sub

Sub fs_CalCom(ByVal r_int_Saltot As Integer, ByVal r_int_NumDeu As Integer, ByVal r_int_PrvCom As Integer, ByVal r_int_Cocoga As Integer, ByVal r_int_CoSnGa As Integer)
      
   'If g_rst_Princi!SITCRE = 1 Then
      If g_rst_Princi!CLAPRV = 0 Then
         l_str_Evalua(r_int_NumDeu) = l_str_Evalua(r_int_NumDeu) + 1
         
         If g_rst_Princi!TIPMON = 1 Then
            l_str_Evalua(r_int_PrvCom - 54) = l_str_Evalua(r_int_PrvCom - 54) + g_rst_Princi!PRVGEN
            l_str_Evalua(r_int_PrvCom) = l_str_Evalua(r_int_PrvCom) + g_rst_Princi!PRVGEN
         Else
            l_str_Evalua(r_int_PrvCom - 54) = l_str_Evalua(r_int_PrvCom - 54) + Format(g_rst_Princi!PRVGEN * g_rst_Princi!TIPCAM, "###,###,##0.00")
            l_str_Evalua(r_int_PrvCom) = l_str_Evalua(r_int_PrvCom) + Format(g_rst_Princi!PRVGEN * g_rst_Princi!TIPCAM, "###,###,##0.00")
         End If
                
         If g_rst_Princi!TIPMON = 1 Then
            l_str_Evalua(r_int_Saltot) = l_str_Evalua(r_int_Saltot) + (g_rst_Princi!SALCAP)
            l_str_Evalua(r_int_Saltot + 54) = l_str_Evalua(r_int_Saltot + 54) + (g_rst_Princi!SALCAP)
         Else
            l_str_Evalua(r_int_Saltot) = l_str_Evalua(r_int_Saltot) + Format((g_rst_Princi!SALCAP) * g_rst_Princi!TIPCAM, "###,###,##0.00")
            l_str_Evalua(r_int_Saltot + 54) = l_str_Evalua(r_int_Saltot + 54) + Format((g_rst_Princi!SALCAP) * g_rst_Princi!TIPCAM, "###,###,##0.00")
         End If
         
         If g_rst_Princi!TIPGAR = 1 Or g_rst_Princi!TIPGAR = 2 Then
            If g_rst_Princi!TIPMON = 1 Then
               l_str_Evalua(r_int_Cocoga) = l_str_Evalua(r_int_Cocoga) + g_rst_Princi!SALCAP
               l_str_Evalua(r_int_Cocoga + 5) = l_str_Evalua(r_int_Cocoga + 5) + g_rst_Princi!SALCAP
            Else
               l_str_Evalua(r_int_Cocoga) = l_str_Evalua(r_int_Cocoga) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
               l_str_Evalua(r_int_Cocoga + 5) = l_str_Evalua(r_int_Cocoga + 5) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
            End If
         Else
            If g_rst_Princi!TIPMON = 1 Then
               l_str_Evalua(r_int_CoSnGa) = l_str_Evalua(r_int_CoSnGa) + g_rst_Princi!SALCAP
               l_str_Evalua(r_int_CoSnGa + 5) = l_str_Evalua(r_int_CoSnGa + 5) + g_rst_Princi!SALCAP
            Else
               l_str_Evalua(r_int_CoSnGa) = l_str_Evalua(r_int_CoSnGa) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
               l_str_Evalua(r_int_CoSnGa + 5) = l_str_Evalua(r_int_CoSnGa + 5) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
            End If
         End If
      End If
   'ElseIf g_rst_Princi!SITCRE = 5 Then
      If g_rst_Princi!CLAPRV = 1 Then
         l_str_Evalua(r_int_NumDeu + 1) = l_str_Evalua(r_int_NumDeu + 1) + 1
         l_str_Evalua(r_int_PrvCom + 1) = l_str_Evalua(r_int_PrvCom + 1) + g_rst_Princi!PRVESP
                                  
         If g_rst_Princi!TIPMON = 1 Then '1
            l_str_Evalua(r_int_Saltot) = l_str_Evalua(r_int_Saltot) + (g_rst_Princi!SALCAP)
         Else
            l_str_Evalua(r_int_Saltot) = l_str_Evalua(r_int_Saltot) + Format((g_rst_Princi!SALCAP) * g_rst_Princi!TIPCAM, "###,###,##0.00")
         End If
          
         If g_rst_Princi!TIPGAR = 1 Or g_rst_Princi!TIPGAR = 2 Then
            If g_rst_Princi!TIPMON = 1 Then
               l_str_Evalua(r_int_Cocoga + 1) = l_str_Evalua(r_int_Cocoga + 1) + g_rst_Princi!SALCAP
               l_str_Evalua(r_int_Cocoga + 5) = l_str_Evalua(r_int_Cocoga + 5) + g_rst_Princi!SALCAP
            Else
               l_str_Evalua(r_int_Cocoga + 1) = l_str_Evalua(r_int_Cocoga + 1) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
               l_str_Evalua(r_int_Cocoga + 5) = l_str_Evalua(r_int_Cocoga + 5) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
            End If
         Else
            If g_rst_Princi!TIPMON = 1 Then
               l_str_Evalua(r_int_CoSnGa + 1) = l_str_Evalua(r_int_CoSnGa + 1) + g_rst_Princi!SALCAP
               l_str_Evalua(r_int_CoSnGa + 5) = l_str_Evalua(r_int_CoSnGa + 5) + g_rst_Princi!SALCAP
            Else
               l_str_Evalua(r_int_CoSnGa + 1) = l_str_Evalua(r_int_CoSnGa + 1) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
               l_str_Evalua(r_int_CoSnGa + 5) = l_str_Evalua(r_int_CoSnGa + 5) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
            End If
         End If
          
      ElseIf g_rst_Princi!CLAPRV = 2 Then
         l_str_Evalua(r_int_NumDeu + 2) = l_str_Evalua(r_int_NumDeu + 2) + 1
         l_str_Evalua(r_int_PrvCom + 2) = l_str_Evalua(r_int_PrvCom + 2) + g_rst_Princi!PRVESP
                  
         If g_rst_Princi!TIPMON = 1 Then
            l_str_Evalua(r_int_Saltot) = l_str_Evalua(r_int_Saltot) + (g_rst_Princi!SALCAP)
         Else
            l_str_Evalua(r_int_Saltot) = l_str_Evalua(r_int_Saltot) + Format((g_rst_Princi!SALCAP) * g_rst_Princi!TIPCAM, "###,###,##0.00")
         End If
            
         If g_rst_Princi!TIPGAR = 1 Or g_rst_Princi!TIPGAR = 2 Then
            If g_rst_Princi!TIPMON = 1 Then
               l_str_Evalua(r_int_Cocoga + 2) = l_str_Evalua(r_int_Cocoga + 2) + g_rst_Princi!SALCAP
               l_str_Evalua(r_int_Cocoga + 5) = l_str_Evalua(r_int_Cocoga + 5) + g_rst_Princi!SALCAP
            Else
               l_str_Evalua(r_int_Cocoga + 2) = l_str_Evalua(r_int_Cocoga + 2) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
               l_str_Evalua(r_int_Cocoga + 5) = l_str_Evalua(r_int_Cocoga + 5) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
            End If
         Else
            If g_rst_Princi!TIPMON = 1 Then
               l_str_Evalua(r_int_CoSnGa + 2) = l_str_Evalua(r_int_CoSnGa + 2) + g_rst_Princi!SALCAP
               l_str_Evalua(r_int_CoSnGa + 5) = l_str_Evalua(r_int_CoSnGa + 5) + g_rst_Princi!SALCAP
            Else
               l_str_Evalua(r_int_CoSnGa + 2) = l_str_Evalua(r_int_CoSnGa + 2) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
               l_str_Evalua(r_int_CoSnGa + 5) = l_str_Evalua(r_int_CoSnGa + 5) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
            End If
         End If
       
      ElseIf g_rst_Princi!CLAPRV = 3 Then
         l_str_Evalua(r_int_NumDeu + 3) = l_str_Evalua(r_int_NumDeu + 3) + 1
         l_str_Evalua(r_int_PrvCom + 3) = l_str_Evalua(r_int_PrvCom + 3) + g_rst_Princi!PRVESP
         
         If g_rst_Princi!TIPMON = 1 Then
            l_str_Evalua(r_int_Saltot) = l_str_Evalua(r_int_Saltot) + (g_rst_Princi!SALCAP)
         Else
            l_str_Evalua(r_int_Saltot) = l_str_Evalua(r_int_Saltot) + Format((g_rst_Princi!SALCAP) * g_rst_Princi!TIPCAM, "###,###,##0.00")
         End If
         
         If g_rst_Princi!TIPGAR = 1 Or g_rst_Princi!TIPGAR = 2 Then
            If g_rst_Princi!HIPCIE_TIPMON = 1 Then
               l_str_Evalua(r_int_Cocoga + 3) = l_str_Evalua(r_int_Cocoga + 3) + g_rst_Princi!SALCAP
               l_str_Evalua(r_int_Cocoga + 5) = l_str_Evalua(r_int_Cocoga + 5) + g_rst_Princi!SALCAP
            Else
               l_str_Evalua(r_int_Cocoga + 3) = l_str_Evalua(r_int_Cocoga + 3) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
               l_str_Evalua(r_int_Cocoga + 5) = l_str_Evalua(r_int_Cocoga + 5) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
            End If
         Else
            If g_rst_Princi!TIPMON = 1 Then
               l_str_Evalua(r_int_CoSnGa + 3) = l_str_Evalua(r_int_CoSnGa + 3) + g_rst_Princi!SALCAP
               l_str_Evalua(r_int_CoSnGa + 5) = l_str_Evalua(r_int_CoSnGa + 5) + g_rst_Princi!SALCAP
            Else
               l_str_Evalua(r_int_CoSnGa + 3) = l_str_Evalua(r_int_CoSnGa + 3) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
               l_str_Evalua(r_int_CoSnGa + 5) = l_str_Evalua(r_int_CoSnGa + 5) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
            End If
         End If
         
      ElseIf g_rst_Princi!CLAPRV = 4 Then
         l_str_Evalua(r_int_NumDeu + 4) = l_str_Evalua(r_int_NumDeu + 4) + 1
         l_str_Evalua(r_int_PrvCom + 4) = l_str_Evalua(r_int_PrvCom + 4) + g_rst_Princi!PRVESP
         
         If g_rst_Princi!TIPMON = 1 Then
            l_str_Evalua(r_int_Saltot) = l_str_Evalua(r_int_Saltot) + (g_rst_Princi!SALCAP)
         Else
            l_str_Evalua(r_int_Saltot) = l_str_Evalua(r_int_Saltot) + Format((g_rst_Princi!SALCAP) * g_rst_Princi!TIPCAM, "###,###,##0.00")
         End If
         
         If g_rst_Princi!TIPGAR = 1 Or g_rst_Princi!TIPGAR = 2 Then
            If g_rst_Princi!TIPMON = 1 Then
               l_str_Evalua(r_int_Cocoga + 4) = l_str_Evalua(r_int_Cocoga + 4) + g_rst_Princi!HIPCIE_SALCAP
               l_str_Evalua(r_int_Cocoga + 5) = l_str_Evalua(r_int_Cocoga + 5) + g_rst_Princi!HIPCIE_SALCAP
            Else
               l_str_Evalua(r_int_Cocoga + 4) = l_str_Evalua(r_int_Cocoga + 4) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
               l_str_Evalua(r_int_Cocoga + 5) = l_str_Evalua(r_int_Cocoga + 5) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
            End If
         Else
            If g_rst_Princi!TIPMON = 1 Then
               l_str_Evalua(r_int_CoSnGa + 4) = l_str_Evalua(r_int_CoSnGa + 4) + g_rst_Princi!SALCAP
               l_str_Evalua(r_int_CoSnGa + 5) = l_str_Evalua(r_int_CoSnGa + 5) + g_rst_Princi!SALCAP
            Else
               l_str_Evalua(r_int_CoSnGa + 4) = l_str_Evalua(r_int_CoSnGa + 4) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
               l_str_Evalua(r_int_CoSnGa + 5) = l_str_Evalua(r_int_CoSnGa + 5) + Format(g_rst_Princi!SALCAP * g_rst_Princi!TIPCAM, "###,###,##0.00")
            End If
         End If
      End If
   'End If
    
   'If g_rst_Princi!TIPMON = 1 Then
   '   l_str_Evalua(r_int_Saltot + 5) = l_str_Evalua(r_int_Saltot + 5) + (g_rst_Princi!SALCAP)
   'Else
   '   l_str_Evalua(r_int_Saltot + 5) = l_str_Evalua(r_int_Saltot + 5) + Format((g_rst_Princi!SALCAP) * g_rst_Princi!TIPCAM, "###,###,##0.00")
   'End If
      
   'l_str_Evalua(r_int_NumDeu + 5) = l_str_Evalua(r_int_NumDeu + 5) + 1
   'l_str_Evalua(r_int_PrvCom + 5) = l_str_Evalua(r_int_PrvCom + 5) + g_rst_Princi!PRVGEN + g_rst_Princi!PRVESP

End Sub

Private Sub fs_GenExc_Det()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_str_PerMes     As String
   Dim r_str_PerAno     As String
   Dim r_dbl_TipCam     As Double
   Dim l_lngper         As String
 
   g_str_Parame = "SELECT HIPCIE_NUMOPE, HIPCIE_TIPGAR, HIPCIE_CLAPRV, HIPCIE_PRVGEN, HIPCIE_PRVESP, HIPCIE_PRVCIC, HIPCIE_SITCRE, HIPCIE_SALCAP, HIPCIE_CAPVIG, HIPCIE_SALCON, HIPCIE_TIPCAM, HIPCIE_TIPMON, HIPCIE_CODPRD, HIPCIE_FECDES, HIPCIE_INTDIF FROM CRE_HIPCIE H WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_NUMOPE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   
   r_obj_Excel.Sheets(1).Name = "HIPOTECARIOS"
   
   With r_obj_Excel.Sheets(1)
   
      '.Pictures.Insert ("\\Server_micasita\COMUN\FIRMAS\Micasita_Especialistas.gif")
      '.DrawingObjects(1).Left = 20
      '.DrawingObjects(1).Top = 20
      
      '.Range(.Cells(1, 36), .Cells(2, 36)).HorizontalAlignment = xlHAlignRight
      '.Cells(1, 36) = "Dpto. de Tecnología e Informática"
      '.Cells(2, 36) = "Desarrollo de Sistemas"
       
      '.Range(.Cells(5, 18), .Cells(5, 18)).HorizontalAlignment = xlHAlignCenter
      '.Range(.Cells(5, 18), .Cells(5, 18)).Font.Bold = True
      '.Range(.Cells(5, 18), .Cells(5, 18)).Font.Underline = xlUnderlineStyleSingle
      '.Cells(5, 18) = "Créditos Hipotecarios"
   
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO DE OPERACION"
      .Cells(1, 3) = "TIPO DE GARANTIA"
      .Cells(1, 4) = "CLASIFICACION"
      .Cells(1, 5) = "PROVISION GENERICA"
      .Cells(1, 6) = "PROVISION ESPECIFICA"
      .Cells(1, 7) = "PROVISION PRO-CICLICA"
      .Cells(1, 8) = "SITUACION DEL CREDITO"
      .Cells(1, 9) = "SALDO CAPITAL"
      .Cells(1, 10) = "SALDO CONCESIONAL"
      .Cells(1, 11) = "CAPITAL VIGENTE"
      .Cells(1, 12) = "INTERES DIFERIDO"
      .Cells(1, 13) = "TIPO DE CAMBIO"
      .Cells(1, 14) = "TIPO DE MONEDA"
      .Cells(1, 15) = "PRODUCTO"
      .Cells(1, 16) = "FECHA DE DESEMBOLSO"
       
      .Range(.Cells(1, 1), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 16)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 16)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(1, 1), .Cells(1, 16)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 16)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 16)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 16)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 16)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Columns("A").ColumnWidth = 5
      
      .Columns("B").ColumnWidth = 19
      .Columns("B").HorizontalAlignment = xlHAlignCenter
            
      .Columns("C").ColumnWidth = 30
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 14
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 20
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 21
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 23
      .Columns("G").HorizontalAlignment = xlHAlignCenter
            
      .Columns("H").ColumnWidth = 22
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 14
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 20
      .Columns("J").HorizontalAlignment = xlHAlignCenter
            
      .Columns("K").ColumnWidth = 16
      .Columns("K").HorizontalAlignment = xlHAlignCenter
            
      .Columns("L").ColumnWidth = 17
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      
      .Columns("M").ColumnWidth = 15
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      
      .Columns("N").ColumnWidth = 21
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      
      .Columns("O").ColumnWidth = 48
      .Columns("O").HorizontalAlignment = xlHAlignCenter
            
      .Columns("P").ColumnWidth = 22
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("P").NumberFormat = "@"
            
      
   
   g_rst_Princi.MoveFirst
      
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
         
      .Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      .Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!HIPCIE_NUMOPE)
      .Cells(r_int_ConVer, 3) = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPCIE_TIPGAR))
      
      If g_rst_Princi!HIPCIE_CLAPRV = 0 Then
         .Cells(r_int_ConVer, 4) = "NORMAL"
      ElseIf g_rst_Princi!HIPCIE_CLAPRV = 1 Then
         .Cells(r_int_ConVer, 4) = "CPP"
      ElseIf g_rst_Princi!HIPCIE_CLAPRV = 2 Then
         .Cells(r_int_ConVer, 4) = "DEFICIENTE"
      ElseIf g_rst_Princi!HIPCIE_CLAPRV = 3 Then
         .Cells(r_int_ConVer, 4) = "DUDOSO"
      ElseIf g_rst_Princi!HIPCIE_CLAPRV = 4 Then
         .Cells(r_int_ConVer, 4) = "PERDIDA"
      End If
      
      .Cells(r_int_ConVer, 5) = Format(g_rst_Princi!HIPCIE_PRVGEN, "###,###,##0.00")
      .Cells(r_int_ConVer, 6) = Format(g_rst_Princi!HIPCIE_PRVESP, "###,###,##0.00")
      .Cells(r_int_ConVer, 7) = Format(g_rst_Princi!HIPCIE_PRVCIC, "###,###,##0.00")
      .Cells(r_int_ConVer, 8) = IIf(Trim(g_rst_Princi!HIPCIE_SITCRE) = 1, "VIGENTE", "VENCIDO")
      .Cells(r_int_ConVer, 9) = Format(g_rst_Princi!HIPCIE_SALCAP, "###,###,##0.00")
      .Cells(r_int_ConVer, 10) = Format(g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
      .Cells(r_int_ConVer, 11) = Format(g_rst_Princi!HIPCIE_CAPVIG, "###,###,##0.00")
      .Cells(r_int_ConVer, 12) = Format(g_rst_Princi!HIPCIE_INTDIF, "###,###,##0.00")
      .Cells(r_int_ConVer, 13) = Format(g_rst_Princi!HIPCIE_TIPCAM, "##0.000")
      .Cells(r_int_ConVer, 14) = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPCIE_TIPMON))
      .Cells(r_int_ConVer, 15) = moddat_gf_Consulta_Produc(g_rst_Princi!HIPCIE_CODPRD)
      .Cells(r_int_ConVer, 16) = "" & gf_FormatoFecha(g_rst_Princi!HIPCIE_FECDES)
      
      r_int_ConVer = r_int_ConVer + 1
      
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   End With
     
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   g_str_Parame = "SELECT MAX(COMCIE_NDOCLI) AS NDOCLI, MAX(COMCIE_NUECRE) AS NUECRE, MAX(COMCIE_TIPGAR) AS TIPGAR, MAX(COMCIE_CLAPRV) AS CLAPRV, SUM(COMCIE_PRVGEN) AS PRVGEN, SUM(COMCIE_PRVESP) AS PRVESP, SUM(COMCIE_PRVCIC) AS PRVCIC, MAX(COMCIE_SITCRE) AS SITCRE, SUM(COMCIE_SALCAP) AS SALCAP, MAX(COMCIE_TIPCAM) AS TIPCAM, MAX(COMCIE_TIPMON) AS TIPMON FROM CRE_COMCIE H WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & ipp_PerAno.Text & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
   g_str_Parame = g_str_Parame & "GROUP BY COMCIE_NDOCLI "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   r_obj_Excel.Sheets(2).Name = "COMERCIALES"
   
   With r_obj_Excel.Sheets(2)
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NUMERO DE DOCUMENTO"
      .Cells(1, 3) = "CLASE DE PRODUCTO"
      .Cells(1, 4) = "TIPO DE GARANTIA"
      .Cells(1, 5) = "CLASIFICACION"
      .Cells(1, 6) = "PROVISION GENERICA"
      .Cells(1, 7) = "PROVISION ESPECIFICA"
      .Cells(1, 8) = "PROVISION PRO-CICLICA"
      .Cells(1, 9) = "SITUACION DEL CREDITO"
      .Cells(1, 10) = "SALDO CAPITAL"
      .Cells(1, 11) = "TIPO DE CAMBIO"
      .Cells(1, 12) = "TIPO DE MONEDA"
   
      .Range(.Cells(1, 1), .Cells(1, 12)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 12)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      
      .Columns("B").ColumnWidth = 24
      .Columns("B").HorizontalAlignment = xlHAlignCenter
            
      .Columns("C").ColumnWidth = 19
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      
      .Columns("D").ColumnWidth = 18
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      
      .Columns("E").ColumnWidth = 14
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      
      .Columns("F").ColumnWidth = 20
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      
      .Columns("G").ColumnWidth = 21
      .Columns("G").HorizontalAlignment = xlHAlignCenter
            
      .Columns("H").ColumnWidth = 23
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Columns("I").ColumnWidth = 22
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 24
      .Columns("J").HorizontalAlignment = xlHAlignCenter
            
      .Columns("K").ColumnWidth = 15
      .Columns("K").HorizontalAlignment = xlHAlignCenter
            
      .Columns("L").ColumnWidth = 21
      .Columns("L").HorizontalAlignment = xlHAlignCenter
         
      g_rst_Princi.MoveFirst
        
      r_int_ConVer = 2
      
      Do While Not g_rst_Princi.EOF
            
         .Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         .Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!NDOCLI)
         .Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!NUECRE)
         .Cells(r_int_ConVer, 4) = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!TIPGAR))
         .Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!CLAPRV)
         .Cells(r_int_ConVer, 6) = Format(g_rst_Princi!PRVGEN, "###,###,##0.00")
         .Cells(r_int_ConVer, 7) = Format(g_rst_Princi!PRVESP, "###,###,##0.00")
         .Cells(r_int_ConVer, 8) = Format(g_rst_Princi!PRVCIC, "###,###,##0.00")
         .Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!SITCRE)
         .Cells(r_int_ConVer, 10) = Format(g_rst_Princi!SALCAP, "###,###,##0.00")
         .Cells(r_int_ConVer, 11) = Format(g_rst_Princi!TIPCAM, "##0.000")
         .Cells(r_int_ConVer, 12) = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!TIPMON))
                                 
         r_int_ConVer = r_int_ConVer + 1
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
End Sub



