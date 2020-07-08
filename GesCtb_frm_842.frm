VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptSun_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   2775
   ClientLeft      =   5895
   ClientTop       =   8055
   ClientWidth     =   5880
   Icon            =   "GesCtb_frm_842.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2835
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5955
      _Version        =   65536
      _ExtentX        =   10504
      _ExtentY        =   5001
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
         TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   10
            Top             =   315
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F05.1. - Libro Diario"
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
            Picture         =   "GesCtb_frm_842.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   11
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
         Begin VB.CommandButton cmd_ExpTxt 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_842.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Generar Archivo de Texto"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_842.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_842.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5220
            Picture         =   "GesCtb_frm_842.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   2850
            Top             =   30
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
         Height          =   1275
         Left            =   30
         TabIndex        =   12
         Top             =   1470
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
         _ExtentY        =   2249
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2805
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   450
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
            Left            =   1560
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
            Caption         =   "Tipo de Moneda:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   150
            Width           =   1305
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   150
            TabIndex        =   14
            Top             =   510
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   150
            TabIndex        =   13
            Top             =   870
            Width           =   885
         End
      End
   End
End
Attribute VB_Name = "frm_RptSun_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_str_Evalua()        As String

Private Sub cmd_ExpExc_Click()
   'valida
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "La fecha de inicio no puede ser mayor a la fecha de final.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   'confirma
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'procesa
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpTxt_Click()
   'valida
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "La fecha de inicio no puede ser mayor a la fecha de final.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If

   'confirma
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'procesa
   Screen.MousePointer = 11
   Call fs_GenArchivo1 'Call fs_GenTxt
   Call fs_GenArchivo2
   Screen.MousePointer = 0
End Sub
Private Sub fs_GenArchivo1()
Dim r_int_PerAno  As Integer
Dim r_int_PerMes  As Integer
Dim r_str_separa  As String
Dim r_int_NumRes  As Integer
Dim r_str_NomRes  As String
Dim r_str_FecRpt  As String
Dim r_str_IdeLib  As String
Dim r_str_NumRuc  As String
Dim r_str_Cadena  As String

   r_int_PerAno = Year(ipp_FecIni.Text)
   r_int_PerMes = Month(ipp_FecIni.Text)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT EMPGRP_NUMRUC FROM MNT_EMPGRP "
   g_str_Parame = g_str_Parame & "  WHERE EMPGRP_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      r_str_NumRuc = Trim(g_rst_Princi!EMPGRP_NUMRUC)
   End If
   
   g_str_Parame = ""
   g_str_Parame = "USP_RPT_LIBRO_ELECTR ("
   g_str_Parame = g_str_Parame & "" & r_int_PerMes & ", "
   g_str_Parame = g_str_Parame & "" & r_int_PerAno & ","
   g_str_Parame = g_str_Parame & "'REPORTE LIBRO ELECTRONICO', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "',1)"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      moddat_g_int_CntErr = moddat_g_int_CntErr + 1
   Else
      moddat_g_int_FlgGOK = True
   End If
   
   If moddat_g_int_CntErr = 6 Then
      If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Exit Sub
      Else
          moddat_g_int_CntErr = 0
      End If
   End If
  
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Screen.MousePointer = 11
      g_rst_Princi.MoveFirst

      r_str_separa = "|"
      r_str_FecRpt = r_int_PerAno & Format(r_int_PerMes, "00") & "00"
      r_str_IdeLib = "050100"
      r_str_NomRes = moddat_g_str_RutLoc & "\LE" & r_str_NumRuc & r_str_FecRpt & r_str_IdeLib & "001111" & ".txt"
   
      'Creando Archivo
      r_int_NumRes = FreeFile
      Open r_str_NomRes For Output As r_int_NumRes
   
      Do While Not g_rst_Princi.EOF
         r_str_Cadena = ""
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!C1) & r_str_separa & Trim(g_rst_Princi!C2) & r_str_separa & Trim(g_rst_Princi!C3) & r_str_separa
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!C4) & r_str_separa & Trim(g_rst_Princi!C5) & r_str_separa & Trim(g_rst_Princi!C6) & r_str_separa
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!C7) & r_str_separa & Trim(g_rst_Princi!C8) & r_str_separa & Trim(g_rst_Princi!C9) & r_str_separa
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!C10) & r_str_separa & Trim(g_rst_Princi!C11) & r_str_separa & Trim(g_rst_Princi!C12) & r_str_separa
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!C13) & r_str_separa & Trim(g_rst_Princi!C14) & r_str_separa & Trim(g_rst_Princi!C15) & r_str_separa
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!C16) & r_str_separa & Trim(g_rst_Princi!C17) & r_str_separa & Format(g_rst_Princi!C18, "0.00") & r_str_separa
         r_str_Cadena = r_str_Cadena & Format(g_rst_Princi!C19, "0.00") & r_str_separa & Trim(g_rst_Princi!C20) & r_str_separa & Trim(g_rst_Princi!C21) & r_str_separa
         
         Print #r_int_NumRes, r_str_Cadena
         
         g_rst_Princi.MoveNext
      Loop
   End If
   'Cerrando Archivo Resumen
   Close r_int_NumRes
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   MsgBox "Archivo Generado Exitosamente en :  " & moddat_g_str_RutLoc & "\" & "LE" & r_str_NumRuc & r_str_FecRpt & r_str_IdeLib & "001111" & ".txt", vbExclamation, modgen_g_str_NomPlt
   
End Sub
Private Sub fs_GenArchivo2()
Dim r_int_PerAno  As Integer
Dim r_int_PerMes  As Integer
Dim r_str_separa  As String
Dim r_int_NumRes  As Integer
Dim r_str_NomRes  As String
Dim r_str_FecRpt  As String
Dim r_str_IdeLib  As String
Dim r_str_NumRuc  As String
Dim r_str_Cadena  As String

   r_int_PerAno = Year(ipp_FecIni.Text)
   r_int_PerMes = Month(ipp_FecIni.Text)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT EMPGRP_NUMRUC FROM MNT_EMPGRP "
   g_str_Parame = g_str_Parame & "  WHERE EMPGRP_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      r_str_NumRuc = Trim(g_rst_Princi!EMPGRP_NUMRUC)
   End If
   
   g_str_Parame = ""
   g_str_Parame = "USP_RPT_LIBRO_ELECTR ("
   g_str_Parame = g_str_Parame & "" & r_int_PerMes & ", "
   g_str_Parame = g_str_Parame & "" & r_int_PerAno & ","
   g_str_Parame = g_str_Parame & "'REPORTE LIBRO ELECTRONICO', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "',2)"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      moddat_g_int_CntErr = moddat_g_int_CntErr + 1
   Else
      moddat_g_int_FlgGOK = True
   End If
   
   If moddat_g_int_CntErr = 6 Then
      If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
          Exit Sub
      Else
          moddat_g_int_CntErr = 0
      End If
   End If
  
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Screen.MousePointer = 11
      g_rst_Princi.MoveFirst

      r_str_separa = "|"
      r_str_FecRpt = r_int_PerAno & Format(r_int_PerMes, "00") & "00"
      r_str_IdeLib = "050300"
      r_str_NomRes = moddat_g_str_RutLoc & "\LE" & r_str_NumRuc & r_str_FecRpt & r_str_IdeLib & "001111" & ".txt"
   
      'Creando Archivo
      r_int_NumRes = FreeFile
      Open r_str_NomRes For Output As r_int_NumRes
   
      Do While Not g_rst_Princi.EOF
         r_str_Cadena = ""
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!C1) & r_str_separa & Trim(g_rst_Princi!C2) & r_str_separa & Trim(g_rst_Princi!C3) & r_str_separa
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!C4) & r_str_separa & Trim(g_rst_Princi!C5) & r_str_separa & Trim(g_rst_Princi!C6) & r_str_separa
         r_str_Cadena = r_str_Cadena & Trim(g_rst_Princi!C7) & r_str_separa & Trim(g_rst_Princi!C8) & r_str_separa
                 
         Print #r_int_NumRes, r_str_Cadena
         
         g_rst_Princi.MoveNext
      Loop
   End If
   'Cerrando Archivo Resumen
   Close r_int_NumRes
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   MsgBox "Archivo Generado Exitosamente en :  " & moddat_g_str_RutLoc & "\" & "LE" & r_str_NumRuc & r_str_FecRpt & r_str_IdeLib & "001111" & ".txt", vbExclamation, modgen_g_str_NomPlt
   
End Sub
Private Sub cmd_Imprim_Click()
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
       
   'valida
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "La fecha de inicio no puede ser mayor a la fecha de final.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   'confirma
   If MsgBox("¿Está seguro de Imprimir el Libro Diario?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
       
   'procesa
   Screen.MousePointer = 11
   Call fs_GenExc_ExceL
    
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CTB_LIBDIR"
   Dim r_str As String
   r_str = cmb_TipMon.ItemData(cmb_TipMon.ListIndex)
   If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 2 Then
      crp_Imprim.SelectionFormula = "{CTB_LIBDIR.LIBDIR_FLGMON} = " & cmb_TipMon.ItemData(cmb_TipMon.ListIndex) & " AND "
   Else
      crp_Imprim.SelectionFormula = "{CTB_LIBDIR.LIBDIR_FLGMON} < " & 999 & " AND "
   End If

   r_str_FecIni = Right(ipp_FecIni.Text, 4) & Mid(ipp_FecIni.Text, 4, 2) & Left(ipp_FecIni.Text, 2)
   r_str_FecFin = Right(ipp_FecFin.Text, 4) & Mid(ipp_FecFin.Text, 4, 2) & Left(ipp_FecFin.Text, 2)
   
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CTB_LIBDIR.LIBDIR_FECCTB} >= " & r_str_FecIni & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CTB_LIBDIR.LIBDIR_FECCTB} <= " & r_str_FecFin & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CTB_LIBDIR.LIBDIR_PERANO} = " & Right(ipp_FecIni.Text, 4) & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CTB_LIBDIR.LIBDIR_PERMES} = " & Mid(ipp_FecIni.Text, 4, 2) & " "
   
   crp_Imprim.Formulas(0) = "FecIni = """ & ipp_FecIni.Text & """"
   crp_Imprim.Formulas(1) = "FecFin = """ & ipp_FecFin.Text & """"
   crp_Imprim.Formulas(2) = "TipMon = """ & cmb_TipMon.ItemData(cmb_TipMon.ListIndex) & """"
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_40.RPT"
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   ipp_FecIni = (date - Format(Now, "DD")) + 1
   ipp_FecFin = modsec_gf_Fin_Del_Mes(date) & Mid(date, 3, Len(date))
   Call fs_Carga_TipMon
   
   Call gs_CentraForm(Me)
   Call gs_SetFocus(cmb_TipMon)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Carga_TipMon()
   cmb_TipMon.Clear
   cmb_TipMon.AddItem "MONEDA NACIONAL"
   cmb_TipMon.ItemData(cmb_TipMon.NewIndex) = "1"
   cmb_TipMon.AddItem "MONEDA EXTRANJERA"
   cmb_TipMon.ItemData(cmb_TipMon.NewIndex) = "2"
End Sub

Private Sub fs_GenTxt()
Dim r_str_nomFile  As String
Dim r_str_rutFile  As String
Dim r_int_NumFile  As String
Dim r_str_FecIni   As String
Dim r_str_FecFin   As String
Dim r_str_ImpDeb   As String
Dim r_str_ImpHab   As String
Dim r_str_NumRuc   As String
Dim r_str_PerAno   As String
Dim r_str_PerMes   As String
Dim r_str_intOpe   As String
Dim r_str_libReg   As String
Dim r_str_intMon   As String
Dim r_str_moneda   As String
   
   r_str_NumRuc = Format("20511904162", "00000000000")
   r_str_PerAno = Right(ipp_FecIni.Text, 4)
   r_str_PerMes = Mid(ipp_FecIni.Text, 4, 2)
   r_str_intOpe = 1
   r_str_libReg = 1
   r_str_intMon = IIf(cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 2, 2, 1)
   r_str_FecIni = Right(ipp_FecIni.Text, 4) & Mid(ipp_FecIni.Text, 4, 2) & Left(ipp_FecIni.Text, 2)
   r_str_FecFin = Right(ipp_FecFin.Text, 4) & Mid(ipp_FecFin.Text, 4, 2) & Left(ipp_FecFin.Text, 2)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.LIBDIR_NROLIB, A.LIBDIR_TIPNOT, A.LIBDIR_NROASI, A.LIBDIR_FECCTB, "
   g_str_Parame = g_str_Parame & "       A.LIBDIR_GLOCTB, A.LIBDIR_CODCTA, A.LIBDIR_DENCTA, A.LIBDIR_DEBSOL, "
   g_str_Parame = g_str_Parame & "       A.LIBDIR_HABSOL, A.LIBDIR_DEBDOL, A.LIBDIR_HABDOL, A.SEGUSUCRE, "
   g_str_Parame = g_str_Parame & "       B.EMPGRP_RAZSOC , B.EMPGRP_NUMRUC, A.LIBDIR_FLGMON, A.LIBDIR_NUMITE "
   g_str_Parame = g_str_Parame & "  FROM CTB_LIBDIR A "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_EMPGRP B ON A.LIBDIR_CODEMP = B.EMPGRP_CODIGO "
   g_str_Parame = g_str_Parame & "   AND A.LIBDIR_PERANO = " & Right(ipp_FecIni.Text, 4)
   g_str_Parame = g_str_Parame & "   AND A.LIBDIR_PERMES = " & Mid(ipp_FecIni.Text, 4, 2)
   g_str_Parame = g_str_Parame & "   AND A.LIBDIR_FECCTB >= " & r_str_FecIni
   g_str_Parame = g_str_Parame & "   AND A.LIBDIR_FECCTB <= " & r_str_FecFin
   If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 2 Then
      g_str_Parame = g_str_Parame & " AND A.LIBDIR_FLGMON = " & cmb_TipMon.ItemData(cmb_TipMon.ListIndex)
   Else
      g_str_Parame = g_str_Parame & " AND A.LIBDIR_FLGMON < " & 999
   End If
   g_str_Parame = g_str_Parame & " ORDER BY LIBDIR_FECCTB ASC, LIBDIR_NROLIB ASC, LIBDIR_TIPNOT ASC, LIBDIR_NROASI ASC, A.LIBDIR_NUMITE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No existen datos para el periodo seleccionado.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   r_str_nomFile = "LE" & r_str_NumRuc & r_str_PerAno & r_str_PerMes & "0005010000" & r_str_intOpe & r_str_libReg & r_str_intMon & 1
   r_str_rutFile = moddat_g_str_RutLoc & "\" & r_str_nomFile & ".TXT"
   r_int_NumFile = FreeFile
   Open r_str_rutFile For Output As r_int_NumFile
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_str_ImpDeb = ""
      r_str_ImpHab = ""
      r_str_moneda = ""
      If (cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 2) Then
          r_str_ImpDeb = g_rst_Princi!LIBDIR_DEBDOL
          r_str_ImpHab = g_rst_Princi!LIBDIR_HABDOL
          r_str_moneda = "USD"
      Else
          r_str_ImpDeb = g_rst_Princi!LIBDIR_DEBSOL
          r_str_ImpHab = g_rst_Princi!LIBDIR_HABSOL
          r_str_moneda = "PEN"
      End If
      If (CDbl(r_str_ImpDeb) = 0) Then
          r_str_ImpDeb = "0.00"
      End If
      If (CDbl(r_str_ImpHab) = 0) Then
          r_str_ImpHab = "0.00"
      End If
      Print #r_int_NumFile, Format(ipp_FecIni.Text, "yyyymm00") & "|" & _
                                            Trim(g_rst_Princi!LIBDIR_NROASI) & "|" & _
                                            "M" & g_rst_Princi!LIBDIR_NUMITE & "|" & _
                                            Trim(g_rst_Princi!LIBDIR_CODCTA) & "|" & _
                                            "0" & "|" & "0" & "|" & _
                                            r_str_moneda & "|" & _
                                            "0" & "|" & "0" & "|" & "00" & "|" & "0" & "|" & _
                                            Trim(g_rst_Princi!LIBDIR_CODCTA) & "|" & _
                                            gf_FormatoFecha(g_rst_Princi!LIBDIR_FECCTB) & "|" & _
                                            "|" & _
                                            gf_FormatoFecha(g_rst_Princi!LIBDIR_FECCTB) & "|" & _
                                            Trim(g_rst_Princi!LIBDIR_DENCTA) & "|" & _
                                            "0" & "|" & _
                                            r_str_ImpDeb & "|" & r_str_ImpHab & "|" & _
                                            "" & _
                                            "|" & "1" & "|" & "0" & "|"
        
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   Close r_int_NumFile
   
   MsgBox "Operación terminada con exito en la ruta, " & r_str_rutFile & ".", vbInformation, modgen_g_str_NomPlt
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Long
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_int_Contad     As Integer
Dim r_int_ConAux     As Integer
Dim r_str_AsiCtb     As String
Dim r_str_CtaNta     As String
Dim r_lng_ConTem     As Long
Dim r_str_LibCtb     As String
Dim r_str_GloCtb     As String
Dim r_str_FecCtb     As String
Dim r_dbl_DebSol     As Double
Dim r_dbl_DebDol     As Double
Dim r_dbl_HabSol     As Double
Dim r_dbl_HabDol     As Double
Dim r_dbl_ToDeSo     As Double
Dim r_dbl_ToDeDo     As Double
Dim r_dbl_ToHaSo     As Double
Dim r_dbl_ToHaDo     As Double
Dim r_str_Period     As String
Dim r_str_CtaCtb     As String
Dim r_str_DebHab     As String
Dim r_int_PosIni     As Integer
   
   r_str_FecIni = Right(ipp_FecIni.Text, 4) & Mid(ipp_FecIni.Text, 4, 2) & Left(ipp_FecIni.Text, 2)
   r_str_FecFin = Right(ipp_FecFin.Text, 4) & Mid(ipp_FecFin.Text, 4, 2) & Left(ipp_FecFin.Text, 2)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO "
   g_str_Parame = g_str_Parame & " WHERE ANO = " & Right(ipp_FecIni.Text, 4) & " "
   g_str_Parame = g_str_Parame & "   AND MES = " & Mid(ipp_FecIni.Text, 4, 2) & " "
   g_str_Parame = g_str_Parame & "   AND SUBSTR(FECHA_CNTBL,1,10) BETWEEN TO_DATE('" & Trim(ipp_FecIni.Text) & "','DD/MM/YYYY') AND TO_DATE('" & Trim(ipp_FecFin.Text) & "','DD/MM/YYYY') "
   g_str_Parame = g_str_Parame & " ORDER BY FECHA_CNTBL ASC, NRO_LIBRO ASC, TIPO_NOTA ASC, NRO_ASIENTO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No hay datos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Erase r_str_Evalua
   ReDim r_str_Evalua(0)
   
   Do While Not g_rst_Princi.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO_DET "
      g_str_Parame = g_str_Parame & " WHERE ANO = " & Right(ipp_FecIni.Text, 4) & " "
      g_str_Parame = g_str_Parame & "   AND MES = " & Mid(ipp_FecIni.Text, 4, 2) & " "
      g_str_Parame = g_str_Parame & "   AND NRO_LIBRO = " & g_rst_Princi!NRO_LIBRO & " "
      g_str_Parame = g_str_Parame & "   AND NRO_ASIENTO = " & g_rst_Princi!NRO_ASIENTO & " "
      g_str_Parame = g_str_Parame & " ORDER BY FECHA_CNTBL ASC, NRO_LIBRO ASC, NRO_ASIENTO ASC "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      
         g_rst_Genera.MoveFirst
         Do While Not g_rst_Genera.EOF
            ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + IIf(UBound(r_str_Evalua) = 0, 11, 12))
            r_str_Evalua(UBound(r_str_Evalua) - 11) = Trim(g_rst_Genera!NRO_LIBRO)
            r_str_Evalua(UBound(r_str_Evalua) - 10) = Trim(g_rst_Princi!TIPO_NOTA)
            r_str_Evalua(UBound(r_str_Evalua) - 9) = Trim(g_rst_Genera!NRO_ASIENTO)
            r_str_Evalua(UBound(r_str_Evalua) - 8) = Left(Trim(g_rst_Princi!FECHA_CNTBL), 10)
            If IsNull(Trim(g_rst_Princi!DESC_GLOSA)) Then
               r_str_Evalua(UBound(r_str_Evalua) - 7) = ""
            Else
               r_str_Evalua(UBound(r_str_Evalua) - 7) = Trim(g_rst_Princi!DESC_GLOSA)
            End If
            r_str_Evalua(UBound(r_str_Evalua) - 6) = Trim(g_rst_Genera!CNTA_CTBL)
            If IsNull(g_rst_Genera!DET_GLOSA) Then
               r_str_Evalua(UBound(r_str_Evalua) - 5) = ""
            Else
               r_str_Evalua(UBound(r_str_Evalua) - 5) = Trim(g_rst_Genera!DET_GLOSA)
            End If
            If Trim(g_rst_Genera!FLAG_DEBHAB) = "D" Then
               r_str_Evalua(UBound(r_str_Evalua) - 4) = IIf(IsNull(g_rst_Genera!IMP_MOVSOL) = True, 0, Trim(g_rst_Genera!IMP_MOVSOL))
               r_str_Evalua(UBound(r_str_Evalua) - 3) = "0"
               r_str_Evalua(UBound(r_str_Evalua) - 2) = IIf(IsNull(g_rst_Genera!IMP_MOVDOL) = True, 0, Trim(g_rst_Genera!IMP_MOVDOL))
               r_str_Evalua(UBound(r_str_Evalua) - 1) = "0"
            ElseIf Trim(g_rst_Genera!FLAG_DEBHAB) = "H" Then
               r_str_Evalua(UBound(r_str_Evalua) - 4) = "0"
               r_str_Evalua(UBound(r_str_Evalua) - 3) = IIf(IsNull(g_rst_Genera!IMP_MOVSOL) = True, 0, Trim(g_rst_Genera!IMP_MOVSOL))
               r_str_Evalua(UBound(r_str_Evalua) - 2) = "0"
               r_str_Evalua(UBound(r_str_Evalua) - 1) = IIf(IsNull(g_rst_Genera!IMP_MOVDOL) = True, 0, Trim(g_rst_Genera!IMP_MOVDOL))
            End If
            If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 1 Then
               r_str_Evalua(UBound(r_str_Evalua) - 0) = 1
            Else
               If Mid(Trim(g_rst_Genera!CNTA_CTBL), 3, 1) = 2 Then
                  r_str_Evalua(UBound(r_str_Evalua) - 0) = 2
               End If
            End If
            
            g_rst_Genera.MoveNext
            DoEvents
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
               
      g_rst_Princi.MoveNext
      DoEvents
   Loop
               
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "Libro Diario"
   r_obj_Excel.Visible = True
   With r_obj_Excel.Sheets(1)
      .Cells(1, 1) = "FORMATO 5.1: ""LIBRO DIARIO"""
      .Cells(2, 1) = "(" & Trim(cmb_TipMon.Text) & ")"
      .Cells(4, 1) = "PERIODO: "
      .Cells(4, 4) = "Del " & Trim(ipp_FecIni) & " Al " & Trim(ipp_FecFin)
      .Cells(5, 1) = "RUC: "
      .Cells(5, 4) = "20511904162"
      .Cells(6, 1) = "DENOMINACIÓN O RAZÓN SOCIAL: "
      .Cells(6, 4) = "EDPYME MICASITA S.A."

      .Range(.Cells(1, 1), .Cells(8, 5)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(1, 1), .Cells(8, 1)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 3)).Merge
      .Range(.Cells(2, 1), .Cells(2, 3)).Merge
      .Range(.Cells(4, 1), .Cells(4, 3)).Merge
      .Range(.Cells(5, 1), .Cells(5, 3)).Merge
      .Range(.Cells(6, 1), .Cells(6, 3)).Merge
      .Range(.Cells(4, 4), .Cells(4, 5)).Merge
      .Range(.Cells(5, 4), .Cells(5, 5)).Merge
      .Range(.Cells(6, 4), .Cells(6, 5)).Merge
            
      .Cells(9, 1) = "LIBRO"
      .Cells(9, 2) = "NOTA"
      .Cells(9, 3) = "ASIENTO"
      .Cells(9, 4) = "FECHA"
      .Cells(9, 5) = "GLOSA"
      .Cells(9, 6) = "CUENTA CONTABLE ASOCIADA A LA OPERACIÓN"
      .Cells(10, 6) = "CÓDIGO"
      .Cells(10, 7) = "DENOMINACIÓN"
      
      If Trim(cmb_TipMon.ListIndex) = 2 Then
         .Cells(9, 8) = "MOVIMIENTO (US$)"
      Else
         .Cells(9, 8) = "MOVIMIENTO (S/.)"
      End If
      
      .Cells(10, 8) = "DEBE"
      .Cells(10, 9) = "HABER"
      If Trim(cmb_TipMon.ListIndex) = 2 Then
         .Cells(9, 10) = "MOVIMIENTO (S/.)"
      Else
         .Cells(9, 10) = "MOVIMIENTO (US$)"
      End If
      .Cells(10, 10) = "DEBE"
      .Cells(10, 11) = "HABER"
       
      .Columns("A").ColumnWidth = 7
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 9
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("B").NumberFormat = "@"
      .Columns("C").ColumnWidth = 11
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 46
      .Columns("F").ColumnWidth = 15
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("F").NumberFormat = "###########0"
      .Columns("G").ColumnWidth = 44
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 15
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,##0.00"
      
      .Range(.Cells(9, 1), .Cells(10, 1)).Merge
      .Range(.Cells(9, 2), .Cells(10, 2)).Merge
      .Range(.Cells(9, 3), .Cells(10, 3)).Merge
      .Range(.Cells(9, 4), .Cells(10, 4)).Merge
      .Range(.Cells(9, 5), .Cells(10, 5)).Merge
      .Range(.Cells(9, 6), .Cells(9, 7)).Merge
      .Range(.Cells(9, 8), .Cells(9, 9)).Merge
      
      r_int_Contad = 11
      .Range(.Cells(9, 10), .Cells(9, 11)).Merge
      
      .Range(.Cells(9, 1), .Cells(10, r_int_Contad)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(10, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(10, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(10, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(10, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(10, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(10, r_int_Contad)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(9, 1), .Cells(10, r_int_Contad)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(9, 1), .Cells(10, r_int_Contad)).WrapText = True
      .Range(.Cells(9, 1), .Cells(10, r_int_Contad)).Font.Bold = True
                  
      r_int_ConVer = 12
      r_lng_ConTem = 0
      r_dbl_DebSol = 0
      r_dbl_HabSol = 0
      r_dbl_DebDol = 0
      r_dbl_HabDol = 0
      r_dbl_ToDeSo = 0
      r_dbl_ToDeDo = 0
      r_dbl_ToHaSo = 0
      r_dbl_ToHaDo = 0
      r_int_ConAux = 0
      
      For r_lng_ConTem = 0 To UBound(r_str_Evalua) Step 12
         
         If r_str_Evalua(r_lng_ConTem + 11) = IIf(cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 3, 2, cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) Then
            If r_int_ConAux = 0 Then
               r_str_LibCtb = r_str_Evalua(r_lng_ConTem + 0)
               r_str_CtaNta = r_str_Evalua(r_lng_ConTem + 1)
               r_str_AsiCtb = r_str_Evalua(r_lng_ConTem + 2)
               r_str_Period = r_str_Evalua(r_lng_ConTem + 3)
               r_int_ConAux = r_int_ConAux + 1
            End If
            
            If r_str_LibCtb = r_str_Evalua(r_lng_ConTem + 0) And r_str_CtaNta = r_str_Evalua(r_lng_ConTem + 1) And r_str_AsiCtb = r_str_Evalua(r_lng_ConTem + 2) And r_str_Period = r_str_Evalua(r_lng_ConTem + 3) Then
               .Cells(r_int_ConVer, 1) = r_str_Evalua(r_lng_ConTem + 0)
               .Cells(r_int_ConVer, 2) = r_str_Evalua(r_lng_ConTem + 1)
               .Cells(r_int_ConVer, 3) = r_str_Evalua(r_lng_ConTem + 2)
               .Cells(r_int_ConVer, 4) = r_str_Evalua(r_lng_ConTem + 3)
               .Cells(r_int_ConVer, 5) = r_str_Evalua(r_lng_ConTem + 4)
               .Cells(r_int_ConVer, 6) = r_str_Evalua(r_lng_ConTem + 5)
               .Cells(r_int_ConVer, 7) = r_str_Evalua(r_lng_ConTem + 6)
               
               If Trim(cmb_TipMon.ListIndex) = 2 Then
                  .Cells(r_int_ConVer, 8) = r_str_Evalua(r_lng_ConTem + 9)
                  .Cells(r_int_ConVer, 9) = r_str_Evalua(r_lng_ConTem + 10)
                  .Cells(r_int_ConVer, 10) = r_str_Evalua(r_lng_ConTem + 7)
                  .Cells(r_int_ConVer, 11) = r_str_Evalua(r_lng_ConTem + 8)
               Else
                  .Cells(r_int_ConVer, 8) = r_str_Evalua(r_lng_ConTem + 7)
                  .Cells(r_int_ConVer, 9) = r_str_Evalua(r_lng_ConTem + 8)
                  .Cells(r_int_ConVer, 10) = r_str_Evalua(r_lng_ConTem + 9)
                  .Cells(r_int_ConVer, 11) = r_str_Evalua(r_lng_ConTem + 10)
               End If
               
               r_dbl_DebSol = r_dbl_DebSol + CDbl(r_str_Evalua(r_lng_ConTem + 7))
               r_dbl_HabSol = r_dbl_HabSol + CDbl(r_str_Evalua(r_lng_ConTem + 8))
               r_dbl_DebDol = r_dbl_DebDol + CDbl(r_str_Evalua(r_lng_ConTem + 9))
               r_dbl_HabDol = r_dbl_HabDol + CDbl(r_str_Evalua(r_lng_ConTem + 10))
               r_dbl_ToDeSo = r_dbl_ToDeSo + CDbl(r_str_Evalua(r_lng_ConTem + 7))
               r_dbl_ToHaSo = r_dbl_ToHaSo + CDbl(r_str_Evalua(r_lng_ConTem + 8))
               r_dbl_ToDeDo = r_dbl_ToDeDo + CDbl(r_str_Evalua(r_lng_ConTem + 9))
               r_dbl_ToHaDo = r_dbl_ToHaDo + CDbl(r_str_Evalua(r_lng_ConTem + 10))
                           
            Else
               .Cells(r_int_ConVer, 7).HorizontalAlignment = xlHAlignRight
               .Range(.Cells(r_int_ConVer, 7), .Cells(r_int_ConVer, r_int_Contad)).Font.Bold = True
               .Cells(r_int_ConVer, 7) = "TOTAL"
               
               If Trim(cmb_TipMon.ListIndex) = 2 Then
                  .Cells(r_int_ConVer, 10) = r_dbl_DebSol
                  .Cells(r_int_ConVer, 11) = r_dbl_HabSol
                  .Cells(r_int_ConVer, 8) = r_dbl_DebDol
                  .Cells(r_int_ConVer, 9) = r_dbl_HabDol
               Else
                  .Cells(r_int_ConVer, 8) = r_dbl_DebSol
                  .Cells(r_int_ConVer, 9) = r_dbl_HabSol
                  .Cells(r_int_ConVer, 10) = r_dbl_DebDol
                  .Cells(r_int_ConVer, 11) = r_dbl_HabDol
               End If
            
               r_int_ConVer = r_int_ConVer + 2
               r_dbl_DebSol = 0
               r_dbl_HabSol = 0
               r_dbl_DebDol = 0
               r_dbl_HabDol = 0
               r_str_LibCtb = r_str_Evalua(r_lng_ConTem + 0)
               r_str_CtaNta = r_str_Evalua(r_lng_ConTem + 1)
               r_str_AsiCtb = r_str_Evalua(r_lng_ConTem + 2)
               r_str_Period = r_str_Evalua(r_lng_ConTem + 3)
               
               .Cells(r_int_ConVer, 1) = r_str_Evalua(r_lng_ConTem + 0)
               .Cells(r_int_ConVer, 2) = r_str_Evalua(r_lng_ConTem + 1)
               .Cells(r_int_ConVer, 3) = r_str_Evalua(r_lng_ConTem + 2)
               .Cells(r_int_ConVer, 4) = r_str_Evalua(r_lng_ConTem + 3)
               .Cells(r_int_ConVer, 5) = r_str_Evalua(r_lng_ConTem + 4)
               .Cells(r_int_ConVer, 6) = r_str_Evalua(r_lng_ConTem + 5)
               .Cells(r_int_ConVer, 7) = r_str_Evalua(r_lng_ConTem + 6)
               
               If Trim(cmb_TipMon.ListIndex) = 2 Then
                  .Cells(r_int_ConVer, 8) = r_str_Evalua(r_lng_ConTem + 9)
                  .Cells(r_int_ConVer, 9) = r_str_Evalua(r_lng_ConTem + 10)
                  .Cells(r_int_ConVer, 10) = r_str_Evalua(r_lng_ConTem + 7)
                  .Cells(r_int_ConVer, 11) = r_str_Evalua(r_lng_ConTem + 8)
               Else
                  .Cells(r_int_ConVer, 8) = r_str_Evalua(r_lng_ConTem + 7)
                  .Cells(r_int_ConVer, 9) = r_str_Evalua(r_lng_ConTem + 8)
                  .Cells(r_int_ConVer, 10) = r_str_Evalua(r_lng_ConTem + 9)
                  .Cells(r_int_ConVer, 11) = r_str_Evalua(r_lng_ConTem + 10)
               End If
               
               r_dbl_DebSol = r_dbl_DebSol + CDbl(r_str_Evalua(r_lng_ConTem + 7))
               r_dbl_HabSol = r_dbl_HabSol + CDbl(r_str_Evalua(r_lng_ConTem + 8))
               r_dbl_DebDol = r_dbl_DebDol + CDbl(r_str_Evalua(r_lng_ConTem + 9))
               r_dbl_HabDol = r_dbl_HabDol + CDbl(r_str_Evalua(r_lng_ConTem + 10))
               r_dbl_ToDeSo = r_dbl_ToDeSo + CDbl(r_str_Evalua(r_lng_ConTem + 7))
               r_dbl_ToHaSo = r_dbl_ToHaSo + CDbl(r_str_Evalua(r_lng_ConTem + 8))
               r_dbl_ToDeDo = r_dbl_ToDeDo + CDbl(r_str_Evalua(r_lng_ConTem + 9))
               r_dbl_ToHaDo = r_dbl_ToHaDo + CDbl(r_str_Evalua(r_lng_ConTem + 10))
            
            End If
            
            r_int_ConVer = r_int_ConVer + 1
         End If
      Next
      
      .Cells(r_int_ConVer, 7).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(r_int_ConVer, 7), .Cells(r_int_ConVer, r_int_Contad)).Font.Bold = True
      .Cells(r_int_ConVer, 7) = "TOTAL"
      
      If Trim(cmb_TipMon.ListIndex) = 2 Then
         .Cells(r_int_ConVer, 10) = r_dbl_DebSol
         .Cells(r_int_ConVer, 11) = r_dbl_HabSol
         .Cells(r_int_ConVer, 8) = r_dbl_DebDol
         .Cells(r_int_ConVer, 9) = r_dbl_HabDol
      Else
         .Cells(r_int_ConVer, 8) = r_dbl_DebSol
         .Cells(r_int_ConVer, 9) = r_dbl_HabSol
         .Cells(r_int_ConVer, 10) = r_dbl_DebDol
         .Cells(r_int_ConVer, 11) = r_dbl_HabDol
      End If
      
      .Cells(r_int_ConVer + 2, 7).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(r_int_ConVer + 2, 7), .Cells(r_int_ConVer + 2, r_int_Contad)).Font.Bold = True
      .Cells(r_int_ConVer + 2, 7) = "TOTAL GENERAL"
      
      If Trim(cmb_TipMon.ListIndex) = 2 Then
         .Cells(r_int_ConVer + 2, 10) = r_dbl_ToDeSo
         .Cells(r_int_ConVer + 2, 11) = r_dbl_ToHaSo
         .Cells(r_int_ConVer + 2, 8) = r_dbl_ToDeDo
         .Cells(r_int_ConVer + 2, 9) = r_dbl_ToHaDo
      Else
         .Cells(r_int_ConVer + 2, 8) = r_dbl_ToDeSo
         .Cells(r_int_ConVer + 2, 9) = r_dbl_ToHaSo
         .Cells(r_int_ConVer + 2, 10) = r_dbl_ToDeDo
         .Cells(r_int_ConVer + 2, 11) = r_dbl_ToHaDo
      End If
            
      .Range(.Cells(1, 1), .Cells(r_int_ConVer + 2, 11)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_ConVer + 2, 11)).Font.Size = 8
   End With
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_ExceL()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_int_Contad     As Integer
Dim r_int_ConAux     As Integer
Dim r_str_AsiCtb     As String
Dim r_str_CtaNta     As String
Dim r_lng_ConTem     As Long
Dim r_str_LibCtb     As String
Dim r_str_GloCtb     As String
Dim r_str_FecCtb     As String
Dim r_dbl_DebSol     As Double
Dim r_dbl_DebDol     As Double
Dim r_dbl_HabSol     As Double
Dim r_dbl_HabDol     As Double
Dim r_dbl_ToDeSo     As Double
Dim r_dbl_ToDeDo     As Double
Dim r_dbl_ToHaSo     As Double
Dim r_dbl_ToHaDo     As Double
Dim r_str_Period     As String
Dim r_str_CtaCtb     As String
Dim r_str_DebHab     As String
Dim r_int_PosIni     As Integer
   
   r_str_FecIni = Right(ipp_FecIni.Text, 4) & Mid(ipp_FecIni.Text, 4, 2) & Left(ipp_FecIni.Text, 2)
   r_str_FecFin = Right(ipp_FecFin.Text, 4) & Mid(ipp_FecFin.Text, 4, 2) & Left(ipp_FecFin.Text, 2)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO "
   g_str_Parame = g_str_Parame & " WHERE ANO = " & Right(ipp_FecIni.Text, 4) & " "
   g_str_Parame = g_str_Parame & "   AND MES = " & Mid(ipp_FecIni.Text, 4, 2) & " "
   g_str_Parame = g_str_Parame & "   AND SUBSTR(FECHA_CNTBL,1,10) BETWEEN TO_DATE('" & Trim(ipp_FecIni.Text) & "','DD/MM/YYYY') AND TO_DATE('" & Trim(ipp_FecFin.Text) & "','DD/MM/YYYY') "
   g_str_Parame = g_str_Parame & " ORDER BY FECHA_CNTBL ASC, NRO_LIBRO ASC, TIPO_NOTA ASC, NRO_ASIENTO ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No hay datos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Erase r_str_Evalua
   ReDim r_str_Evalua(0)
   Do While Not g_rst_Princi.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO_DET "
      g_str_Parame = g_str_Parame & " WHERE ANO = " & Right(ipp_FecIni.Text, 4) & " "
      g_str_Parame = g_str_Parame & "   AND MES = " & Mid(ipp_FecIni.Text, 4, 2) & " "
      g_str_Parame = g_str_Parame & "   AND NRO_LIBRO = " & g_rst_Princi!NRO_LIBRO & " "
      g_str_Parame = g_str_Parame & "   AND NRO_ASIENTO = " & g_rst_Princi!NRO_ASIENTO & " "
      g_str_Parame = g_str_Parame & " ORDER BY FECHA_CNTBL ASC, NRO_LIBRO ASC, NRO_ASIENTO ASC "
           
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
         Do While Not g_rst_Genera.EOF
            ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + IIf(UBound(r_str_Evalua) = 0, 11, 12))
            r_str_Evalua(UBound(r_str_Evalua) - 11) = Trim(g_rst_Genera!NRO_LIBRO)
            r_str_Evalua(UBound(r_str_Evalua) - 10) = Trim(g_rst_Princi!TIPO_NOTA)
            r_str_Evalua(UBound(r_str_Evalua) - 9) = Trim(g_rst_Genera!NRO_ASIENTO)
            r_str_Evalua(UBound(r_str_Evalua) - 8) = Left(Trim(g_rst_Princi!FECHA_CNTBL), 10)
            r_str_Evalua(UBound(r_str_Evalua) - 7) = IIf(IsNull(g_rst_Princi!DESC_GLOSA) = True, "", Trim(g_rst_Princi!DESC_GLOSA))
            r_str_Evalua(UBound(r_str_Evalua) - 6) = Trim(g_rst_Genera!CNTA_CTBL)
            If IsNull(g_rst_Genera!DET_GLOSA) Then
               r_str_Evalua(UBound(r_str_Evalua) - 5) = ""
            Else
               r_str_Evalua(UBound(r_str_Evalua) - 5) = Trim(g_rst_Genera!DET_GLOSA)
            End If
            If Trim(g_rst_Genera!FLAG_DEBHAB) = "D" Then
               r_str_Evalua(UBound(r_str_Evalua) - 4) = IIf(IsNull(g_rst_Genera!IMP_MOVSOL) = True, 0, Trim(g_rst_Genera!IMP_MOVSOL))
               r_str_Evalua(UBound(r_str_Evalua) - 3) = "0"
               r_str_Evalua(UBound(r_str_Evalua) - 2) = IIf(IsNull(g_rst_Genera!IMP_MOVDOL) = True, 0, Trim(g_rst_Genera!IMP_MOVDOL))
               r_str_Evalua(UBound(r_str_Evalua) - 1) = "0"
            ElseIf Trim(g_rst_Genera!FLAG_DEBHAB) = "H" Then
               r_str_Evalua(UBound(r_str_Evalua) - 4) = "0"
               r_str_Evalua(UBound(r_str_Evalua) - 3) = IIf(IsNull(g_rst_Genera!IMP_MOVSOL) = True, 0, Trim(g_rst_Genera!IMP_MOVSOL))
               r_str_Evalua(UBound(r_str_Evalua) - 2) = "0"
               r_str_Evalua(UBound(r_str_Evalua) - 1) = IIf(IsNull(g_rst_Genera!IMP_MOVDOL) = True, 0, Trim(g_rst_Genera!IMP_MOVDOL))
            End If
            If Mid(Trim(g_rst_Genera!CNTA_CTBL), 3, 1) = 1 Then
               r_str_Evalua(UBound(r_str_Evalua) - 0) = 1
            Else
               r_str_Evalua(UBound(r_str_Evalua) - 0) = 2
            End If
            
            g_rst_Genera.MoveNext
            DoEvents
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
               
      g_rst_Princi.MoveNext
      DoEvents
   Loop

   'ELIMINACION DE LA TABLA TEMPORAL
   g_str_Parame = ""
   g_str_Parame = "DELETE FROM CTB_LIBDIR WHERE "
   g_str_Parame = g_str_Parame & "LIBDIR_CODEMP = '000001' AND "
   g_str_Parame = g_str_Parame & "LIBDIR_PERMES = " & Format(Mid(ipp_FecIni.Text, 4, 2), "00") & " AND "
   g_str_Parame = g_str_Parame & "LIBDIR_PERANO = " & Format(Right(ipp_FecIni.Text, 4), "0000") & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   r_int_ConAux = 0
   For r_lng_ConTem = 0 To UBound(r_str_Evalua) Step 12
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CTB_LIBDIR("
      g_str_Parame = g_str_Parame & "LIBDIR_CODEMP, "
      g_str_Parame = g_str_Parame & "LIBDIR_PERMES, "
      g_str_Parame = g_str_Parame & "LIBDIR_PERANO, "
      g_str_Parame = g_str_Parame & "LIBDIR_NUMITE, "
      g_str_Parame = g_str_Parame & "LIBDIR_NROLIB, "
      g_str_Parame = g_str_Parame & "LIBDIR_TIPNOT, "
      g_str_Parame = g_str_Parame & "LIBDIR_NROASI, "
      g_str_Parame = g_str_Parame & "LIBDIR_FECCTB, "
      g_str_Parame = g_str_Parame & "LIBDIR_GLOCTB, "
      g_str_Parame = g_str_Parame & "LIBDIR_CODCTA, "
      g_str_Parame = g_str_Parame & "LIBDIR_DENCTA, "
      g_str_Parame = g_str_Parame & "LIBDIR_DEBSOL, "
      g_str_Parame = g_str_Parame & "LIBDIR_HABSOL, "
      g_str_Parame = g_str_Parame & "LIBDIR_DEBDOL, "
      g_str_Parame = g_str_Parame & "LIBDIR_HABDOL, "
      g_str_Parame = g_str_Parame & "LIBDIR_FLGMON, "
      g_str_Parame = g_str_Parame & "SEGUSUCRE) "
      
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "'000001', "
      g_str_Parame = g_str_Parame & Format(Mid(ipp_FecIni.Text, 4, 2), "00") & ", "
      g_str_Parame = g_str_Parame & Format(Right(ipp_FecIni.Text, 4), "0000") & ", "
      g_str_Parame = g_str_Parame & r_int_ConAux & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_Evalua(r_lng_ConTem + 0) & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Evalua(r_lng_ConTem + 1) & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Evalua(r_lng_ConTem + 2) & "', "
      g_str_Parame = g_str_Parame & Format(r_str_Evalua(r_lng_ConTem + 3), "YYYYMMDD") & ", "
      g_str_Parame = g_str_Parame & "'" & Replace(Trim(r_str_Evalua(r_lng_ConTem + 4)), "'", "") & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Evalua(r_lng_ConTem + 5) & "', "
      g_str_Parame = g_str_Parame & "'" & Replace(Trim(r_str_Evalua(r_lng_ConTem + 6)), "'", "") & "', "
      g_str_Parame = g_str_Parame & Format(r_str_Evalua(r_lng_ConTem + 7), "###########0.00") & ", "
      g_str_Parame = g_str_Parame & Format(r_str_Evalua(r_lng_ConTem + 8), "###########0.00") & ","
      g_str_Parame = g_str_Parame & Format(r_str_Evalua(r_lng_ConTem + 9), "###########0.00") & ", "
      g_str_Parame = g_str_Parame & Format(r_str_Evalua(r_lng_ConTem + 10), "###########0.00") & ","
      g_str_Parame = g_str_Parame & r_str_Evalua(r_lng_ConTem + 11) & ","
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "')"
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      r_int_ConAux = r_int_ConAux + 1
   Next
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Imprim)
   End If
End Sub

