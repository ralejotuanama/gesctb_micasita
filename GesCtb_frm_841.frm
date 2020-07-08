VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptSun_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   2700
   ClientLeft      =   7920
   ClientTop       =   5985
   ClientWidth     =   7215
   Icon            =   "GesCtb_frm_841.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2835
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7215
      _Version        =   65536
      _ExtentX        =   12726
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
         TabIndex        =   7
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
            TabIndex        =   8
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
            TabIndex        =   9
            Top             =   315
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Libro Mayor"
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
            Picture         =   "GesCtb_frm_841.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   10
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
         Begin VB.CommandButton cmd_ExpTxt 
            Height          =   585
            Left            =   1200
            Picture         =   "GesCtb_frm_841.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Generar Archivo de Texto"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6510
            Picture         =   "GesCtb_frm_841.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_841.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   600
            Picture         =   "GesCtb_frm_841.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   5640
            Top             =   120
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
         TabIndex        =   11
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
         Begin VB.TextBox txt_CtaCtb 
            Height          =   345
            Left            =   4110
            TabIndex        =   16
            Top             =   600
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   4110
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   2805
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1170
            TabIndex        =   0
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
            TabIndex        =   1
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
         Begin VB.Label Label4 
            Caption         =   "Nº de Cuenta:"
            Height          =   255
            Left            =   3060
            TabIndex        =   15
            Top             =   660
            Visible         =   0   'False
            Width           =   1215
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
Attribute VB_Name = "frm_RptSun_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_dbl_Evalua(30039)   As Double

Private Sub cmd_ExpExc_Click()
    If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
        MsgBox "La fecha de inicio no puede ser mayor a la fecha de final.", vbInformation, modgen_g_str_NomPlt
        Exit Sub
    End If
    If cmb_TipMon.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de moneda.", vbExclamation, modgen_g_str_NomPlt
        Call gs_SetFocus(txt_CtaCtb)
        Exit Sub
    End If
 
    If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Call fs_GenExc_LMayor
    Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpTxt_Click()
   If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
      MsgBox "La fecha de inicio no puede ser mayor a la fecha de fin.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
 
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
    
   Screen.MousePointer = 11
   Call fs_GenArchivo
   Screen.MousePointer = 0
   
End Sub

Private Sub fs_GenArchivo()
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
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "',3)"

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
      r_str_IdeLib = "060100"
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

Private Sub cmd_Imprim_Click()
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
    If MsgBox("¿Está seguro de Imprimir el Libro Mayor?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If
    
    Screen.MousePointer = 11
    Call fs_Imprimir
    Screen.MousePointer = 0
End Sub

Private Sub fs_GenRep()
Dim r_obj_Excel      As Excel.Application
Dim r_int_nrofil     As Integer
Dim r_int_ConFil     As Integer
Dim r_int_NroIte     As Integer
Dim r_int_NroVal     As Integer
Dim r_str_CtaCtb     As String
Dim r_dbl_TotDeu     As Double
Dim r_dbl_TotAcr     As Double
  
   r_int_nrofil = 1
   r_str_CtaCtb = ""
   r_int_ConFil = 2
   r_dbl_TotDeu = 0
   r_dbl_TotAcr = 0
   r_int_NroIte = 0
   r_int_NroVal = 0
   
   '**********************************
   '*********PROCEDURE****************
   '**********************************
   g_str_Parame = ""
   g_str_Parame = "usp_lm_libro_mayor ("
   g_str_Parame = g_str_Parame & "'" & IIf((cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) = 1, "sol", "dol") & "', "
   g_str_Parame = g_str_Parame & "'L',"
   g_str_Parame = g_str_Parame & "'" & ipp_FecIni.Text & "', "
   g_str_Parame = g_str_Parame & "'" & ipp_FecFin.Text & "', "
   g_str_Parame = g_str_Parame & "'L','LM')"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 2) Then
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
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT CNTA_CTBL, DESC_CNTA, TO_CHAR(FECHA_CNTBL,'DD/MM/YYYY') AS FECHA_CNTBL, "
   g_str_Parame = g_str_Parame & "       NRO_LIBRO, TIPO_NOTA, NRO_ASIENTO, DESC_GLOSA, DEBE, HABER, ITEM "
   g_str_Parame = g_str_Parame & "  FROM LM_LIBRO_MAYOR "
   g_str_Parame = g_str_Parame & " ORDER BY CNTA_CTBL ASC,FECHA_CNTBL ASC, NRO_LIBRO ASC, NRO_ASIENTO ASC"
   
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
  
   With r_obj_Excel.Sheets(1)
      .Columns("A").NumberFormat = "@"
      .Columns("C").NumberFormat = "@"
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").NumberFormat = "###,###,##0.00"
      
      g_rst_Princi.MoveFirst
      r_str_CtaCtb = Trim(g_rst_Princi!CNTA_CTBL)
      
      Do While Not g_rst_Princi.EOF
         If Trim(g_rst_Princi!CNTA_CTBL) <> r_str_CtaCtb Then
            If r_int_ConFil = 54 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Van"
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(r_dbl_TotDeu)
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(r_dbl_TotAcr)
               r_int_nrofil = r_int_nrofil + 1
               
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Vienen"
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(r_dbl_TotDeu)
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(r_dbl_TotAcr)
               r_int_nrofil = r_int_nrofil + 1
               r_int_ConFil = 0
            End If
            
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Saldo Final"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = r_dbl_TotDeu - r_dbl_TotAcr
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = 0
            If r_dbl_TotAcr > r_dbl_TotDeu Then
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = 0
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = r_dbl_TotAcr - r_dbl_TotDeu
            End If
            r_int_nrofil = r_int_nrofil + 1
            r_int_ConFil = r_int_ConFil + 1
            r_dbl_TotDeu = 0
            r_dbl_TotAcr = 0
         End If
         
         If r_int_ConFil = 54 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Van"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(r_dbl_TotDeu)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(r_dbl_TotAcr)
            r_int_nrofil = r_int_nrofil + 1
            
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Vienen"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(r_dbl_TotDeu)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(r_dbl_TotAcr)
            r_int_nrofil = r_int_nrofil + 1
            r_int_ConFil = 0
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 1) = Trim(g_rst_Princi!CNTA_CTBL)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 2) = Trim(g_rst_Princi!DESC_CNTA)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 3) = Trim(g_rst_Princi!FECHA_CNTBL)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 4) = Trim(g_rst_Princi!NRO_LIBRO)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 5) = Trim(g_rst_Princi!TIPO_NOTA)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 6) = Trim(g_rst_Princi!NRO_ASIENTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = Trim(g_rst_Princi!DESC_GLOSA)
         
         If CInt(g_rst_Princi!NRO_LIBRO) = 0 And CInt(g_rst_Princi!NRO_ASIENTO) = 0 And CInt(g_rst_Princi!Item) = 0 Then
            If CDbl(g_rst_Princi!DEBE) > CDbl(g_rst_Princi!HABER) Then
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(g_rst_Princi!DEBE) - CDbl(g_rst_Princi!HABER)
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = 0
               r_dbl_TotDeu = r_dbl_TotDeu + CDbl(g_rst_Princi!DEBE) - CDbl(g_rst_Princi!HABER)
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = 0
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(g_rst_Princi!HABER) - CDbl(g_rst_Princi!DEBE)
               r_dbl_TotAcr = r_dbl_TotAcr + CDbl(g_rst_Princi!HABER) - CDbl(g_rst_Princi!DEBE)
            End If
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(g_rst_Princi!DEBE)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(g_rst_Princi!HABER)
            r_dbl_TotDeu = r_dbl_TotDeu + CDbl(g_rst_Princi!DEBE)
            r_dbl_TotAcr = r_dbl_TotAcr + CDbl(g_rst_Princi!HABER)
         End If
         
         r_int_nrofil = r_int_nrofil + 1
         r_int_ConFil = r_int_ConFil + 1
         r_str_CtaCtb = Trim(g_rst_Princi!CNTA_CTBL)
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Saldo Final"
      r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = r_dbl_TotDeu - r_dbl_TotAcr
      r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = 0
      If r_dbl_TotAcr > r_dbl_TotDeu Then
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = 0
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = r_dbl_TotAcr - r_dbl_TotDeu
      End If
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
    
   With r_obj_Excel.Sheets(1)
      r_obj_Excel.ActiveSheet.Range("A54").Select
      r_obj_Excel.ActiveCell.EntireRow.Insert Shift:=xlDown
      .Cells(54, 8) = "NULL": .Cells(54, 9) = "NULL"
      r_obj_Excel.ActiveCell.EntireRow.Insert Shift:=xlDown
      .Cells(54, 8) = "NULL": .Cells(54, 9) = "NULL"
      r_obj_Excel.ActiveCell.EntireRow.Insert Shift:=xlDown
      .Cells(54, 8) = "NULL": .Cells(54, 9) = "NULL"
      
      'ELIMINACION DE LA TABLA TEMPORAL
      g_str_Parame = ""
      g_str_Parame = "DELETE FROM CTB_LIBMAY"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      'INSERT A LA TABLA CTB_LIBMAY
      r_int_NroVal = r_int_nrofil + 3
      If r_int_NroVal < 54 Then
         r_int_NroVal = r_int_nrofil
      End If
      
      For r_int_NroIte = 1 To r_int_NroVal
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO CTB_LIBMAY("
         g_str_Parame = g_str_Parame & "LIBMAY_CODEMP, "
         g_str_Parame = g_str_Parame & "LIBMAY_CTACTB, "
         g_str_Parame = g_str_Parame & "LIBMAY_DSCNTA, "
         g_str_Parame = g_str_Parame & "LIBMAY_FCNTBL, "
         g_str_Parame = g_str_Parame & "LIBMAY_NOLIBR, "
         g_str_Parame = g_str_Parame & "LIBMAY_TPNOTA, "
         g_str_Parame = g_str_Parame & "LIBMAY_NOASNT, "
         g_str_Parame = g_str_Parame & "LIBMAY_DGLOSA, "
         g_str_Parame = g_str_Parame & "LIBMAY_SALDEB, "
         g_str_Parame = g_str_Parame & "LIBMAY_SALHAB, "
         g_str_Parame = g_str_Parame & "LIBMAY_NUMERO, "
         g_str_Parame = g_str_Parame & "SEGUSUCRE ) "
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'000001', "
         g_str_Parame = g_str_Parame & "'" & Trim(r_obj_Excel.Cells(r_int_NroIte, 1)) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(r_obj_Excel.Cells(r_int_NroIte, 2)) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(r_obj_Excel.Cells(r_int_NroIte, 3)) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(r_obj_Excel.Cells(r_int_NroIte, 4)) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(r_obj_Excel.Cells(r_int_NroIte, 5)) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(r_obj_Excel.Cells(r_int_NroIte, 6)) & "', "
         g_str_Parame = g_str_Parame & "'" & Replace(Trim(r_obj_Excel.Cells(r_int_NroIte, 7)), "'", "") & "', "
         g_str_Parame = g_str_Parame & Trim(r_obj_Excel.Cells(r_int_NroIte, 8)) & ", "
         g_str_Parame = g_str_Parame & Trim(r_obj_Excel.Cells(r_int_NroIte, 9)) & ", "
         g_str_Parame = g_str_Parame & "'" & r_int_NroIte & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "')"
             
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
             Exit Sub
         End If
      Next r_int_NroIte
   End With
   
   r_obj_Excel.Visible = False
   r_obj_Excel.ActiveWorkbook.Close False
   
   r_obj_Excel.Quit
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Imprimir()
    'consulta
    Call fs_GenRep
    
    crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
    crp_Imprim.DataFiles(0) = "CTB_LIBMAY"
    crp_Imprim.Formulas(0) = "FecIni = """ & ipp_FecIni.Text & """"
    crp_Imprim.Formulas(1) = "FecFin = """ & ipp_FecFin.Text & """"
    crp_Imprim.Formulas(2) = "TipMon = """ & cmb_TipMon.ItemData(cmb_TipMon.ListIndex) & """"
  
    crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "ctb_rptsol_41.RPT"
    crp_Imprim.Destination = crptToWindow
    crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Caption = modgen_g_str_NomPlt
    
    ipp_FecIni = (date - Format(Now, "DD")) + 1
    ipp_FecFin = modsec_gf_Fin_Del_Mes(date) & Mid(date, 3, Len(date))
    cmb_TipMon.Clear
    cmb_TipMon.AddItem "MONEDA NACIONAL"
    cmb_TipMon.ItemData(cmb_TipMon.NewIndex) = "1"
    cmb_TipMon.AddItem "MONEDA EXTRANJERA"
    cmb_TipMon.ItemData(cmb_TipMon.NewIndex) = "2"
       
    Call gs_CentraForm(Me)
    Call gs_SetFocus(ipp_FecIni)
    Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_int_Contad     As Integer
Dim r_str_CtaCtb     As String
Dim r_str_CtaNta     As String
Dim r_int_ConTem     As Integer
   
   Erase r_dbl_Evalua
   r_str_FecIni = Right(ipp_FecIni.Text, 4) & Mid(ipp_FecIni.Text, 4, 2) & Left(ipp_FecIni.Text, 2)
   r_str_FecFin = Right(ipp_FecFin.Text, 4) & Mid(ipp_FecFin.Text, 4, 2) & Left(ipp_FecFin.Text, 2)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO_DET  "
   g_str_Parame = g_str_Parame & " WHERE FECHA_CNTBL BETWEEN TO_DATE('" & Trim(ipp_FecIni.Text) & "','DD/MM/YYYY') AND TO_DATE('" & Trim(ipp_FecFin.Text) & "','DD/MM/YYYY') "
   If Trim(cmb_TipMon.ListIndex) = 0 Then
      g_str_Parame = g_str_Parame & "   AND SUBSTR(CNTA_CTBL,3,1) = 1"
   ElseIf Trim(cmb_TipMon.ListIndex) = 1 Or Trim(cmb_TipMon.ListIndex) = 2 Then
      g_str_Parame = g_str_Parame & "   AND SUBSTR(CNTA_CTBL,3,1) = 2"
   End If
   If Trim(txt_CtaCtb.Text) <> "" And Trim(cmb_TipMon.ListIndex) <> 3 Then
      g_str_Parame = g_str_Parame & "   AND TRIM(CNTA_CTBL) = " & Trim(txt_CtaCtb.Text) & " "
   End If
   g_str_Parame = g_str_Parame & " ORDER BY CNTA_CTBL ASC, FECHA_CNTBL ASC, NRO_LIBRO ASC, NRO_ASIENTO ASC  "
   
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
   r_obj_Excel.Sheets(1).Name = "Libro Mayor"
   
   With r_obj_Excel.Sheets(1)
      .Cells(1, 1) = "FORMATO 6.1: ""LIBRO MAYOR"""
      .Cells(2, 1) = "(" & cmb_TipMon.Text & ")"
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
      
      .Cells(9, 1) = "FECHA"
      .Cells(9, 2) = "LIBRO"
      .Cells(9, 3) = "NOTA"
      .Cells(9, 4) = "ASIENTO"
      .Cells(9, 5) = "DESCRIPCIÓN O GLOSA DE LA OPERACIÓN"
      
      If Trim(cmb_TipMon.ListIndex) = 1 Then
         .Cells(9, 6) = "SALDOS Y MOVIMIENTOS (US$)"
      Else
         .Cells(9, 6) = "SALDOS Y MOVIMIENTOS (S/.)"
      End If
      
      .Cells(10, 6) = "DEUDOR"
      .Cells(10, 7) = "ACREEDOR"
      
      If Trim(cmb_TipMon.ListIndex) = 2 Or Trim(cmb_TipMon.ListIndex) = 3 Then
         .Cells(9, 8) = "SALDOS Y MOVIMIENTOS (US$)"
         .Cells(10, 8) = "DEUDOR"
         .Cells(10, 9) = "ACREEDOR"
      End If
      
      .Cells(12, 1) = "CUENTA CONTABLE (1): "
      .Cells(12, 3) = Trim(g_rst_Princi!CNTA_CTBL) & " " & modsec_gf_Buscar_NomCta(g_rst_Princi!CNTA_CTBL)
      .Range(.Cells(12, 1), .Cells(12, 2)).Merge
      .Range(.Cells(12, 3), .Cells(12, 6)).Merge
      .Range(.Cells(12, 1), .Cells(12, 7)).Font.Bold = True
      .Range(.Cells(12, 1), .Cells(12, 7)).HorizontalAlignment = xlHAlignLeft
       
      .Columns("A").ColumnWidth = 12
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("A").NumberFormat = "@"
      .Columns("B").ColumnWidth = 8
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 8
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 55
      .Columns("F").ColumnWidth = 15
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 15
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,##0.00"
      
      .Range(.Cells(9, 1), .Cells(10, 1)).Merge
      .Range(.Cells(9, 2), .Cells(10, 2)).Merge
      .Range(.Cells(9, 3), .Cells(10, 3)).Merge
      .Range(.Cells(9, 4), .Cells(10, 4)).Merge
      .Range(.Cells(9, 5), .Cells(10, 5)).Merge
      .Range(.Cells(9, 6), .Cells(9, 7)).Merge
                  
      If Trim(cmb_TipMon.ListIndex) = 0 Or Trim(cmb_TipMon.ListIndex) = 1 Then
         r_int_Contad = 7
      ElseIf Trim(cmb_TipMon.ListIndex) = 2 Or Trim(cmb_TipMon.ListIndex) = 3 Then
         r_int_Contad = 9
         .Range(.Cells(9, 8), .Cells(9, 9)).Merge
      End If
      
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
      
      g_rst_Princi.MoveFirst
      r_int_ConVer = 13
      r_int_ConTem = 0
      
      r_str_CtaCtb = Trim(g_rst_Princi!CNTA_CTBL)
      r_str_CtaNta = modsec_gf_Buscar_NtaCta(g_rst_Princi!ANO, g_rst_Princi!Mes, g_rst_Princi!NRO_LIBRO, g_rst_Princi!NRO_ASIENTO)
      
      Do While Not g_rst_Princi.EOF
         If Trim(g_rst_Princi!CNTA_CTBL) <> r_str_CtaCtb Then
            .Cells(r_int_ConVer, 5).HorizontalAlignment = xlHAlignRight
            .Range(.Cells(r_int_ConVer, 5), .Cells(r_int_ConVer, r_int_Contad)).Font.Bold = True
            .Cells(r_int_ConVer, 5) = "TOTAL"
            .Cells(r_int_ConVer, 6) = r_dbl_Evalua(r_int_ConTem)
            .Cells(r_int_ConVer, 7) = r_dbl_Evalua(r_int_ConTem + 1)
            
            If Trim(cmb_TipMon.ListIndex) = 2 Or Trim(cmb_TipMon.ListIndex) = 3 Then
               .Cells(r_int_ConVer, 8) = r_dbl_Evalua(r_int_ConTem + 2)
               .Cells(r_int_ConVer, 9) = r_dbl_Evalua(r_int_ConTem + 3)
            End If
            
            .Cells(r_int_ConVer + 2, 1) = "CUENTA CONTABLE (1): "
            .Cells(r_int_ConVer + 2, 3) = Trim(g_rst_Princi!CNTA_CTBL) & " " & modsec_gf_Buscar_NomCta(g_rst_Princi!CNTA_CTBL)
            .Range(.Cells(r_int_ConVer + 2, 1), .Cells(r_int_ConVer + 2, 2)).Merge
            .Range(.Cells(r_int_ConVer + 2, 3), .Cells(r_int_ConVer + 2, 6)).Merge
            .Range(.Cells(r_int_ConVer + 2, 1), .Cells(r_int_ConVer + 2, r_int_Contad)).Font.Bold = True
            .Range(.Cells(r_int_ConVer + 2, 1), .Cells(r_int_ConVer + 2, r_int_Contad)).HorizontalAlignment = xlHAlignLeft
            
            r_str_CtaNta = modsec_gf_Buscar_NtaCta(g_rst_Princi!ANO, g_rst_Princi!Mes, g_rst_Princi!NRO_LIBRO, g_rst_Princi!NRO_ASIENTO)
            r_int_ConTem = r_int_ConTem + 4
            r_int_ConVer = r_int_ConVer + 3
         End If
         
         If IsNull(Trim(g_rst_Princi!CNTA_CTBL)) Then
            r_str_CtaCtb = ""
         Else
            r_str_CtaCtb = Trim(g_rst_Princi!CNTA_CTBL)
         End If
         
         .Cells(r_int_ConVer, 1) = Trim(g_rst_Princi!FECHA_CNTBL)
         .Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!NRO_LIBRO)
         .Cells(r_int_ConVer, 3) = r_str_CtaNta
         .Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!NRO_ASIENTO)
         .Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!DET_GLOSA)
         
         If Trim(cmb_TipMon.ListIndex) = 0 Then
            If Trim(g_rst_Princi!FLAG_DEBHAB) = "D" Then
               If IsNull(g_rst_Princi!IMP_MOVSOL) Then
                  .Cells(r_int_ConVer, 6) = 0
               Else
                  .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!IMP_MOVSOL)
                  r_dbl_Evalua(r_int_ConTem) = r_dbl_Evalua(r_int_ConTem) + Trim(g_rst_Princi!IMP_MOVSOL)
               End If
               .Cells(r_int_ConVer, 7) = 0
               
            ElseIf Trim(g_rst_Princi!FLAG_DEBHAB) = "H" Then
               .Cells(r_int_ConVer, 6) = 0
               If IsNull(g_rst_Princi!IMP_MOVSOL) Then
                  .Cells(r_int_ConVer, 7) = 0
               Else
                  .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!IMP_MOVSOL)
                  r_dbl_Evalua(r_int_ConTem + 1) = r_dbl_Evalua(r_int_ConTem + 1) + Trim(g_rst_Princi!IMP_MOVSOL)
               End If
            End If
            
         ElseIf Trim(cmb_TipMon.ListIndex) = 1 Then
            If Trim(g_rst_Princi!FLAG_DEBHAB) = "D" Then
               If IsNull(g_rst_Princi!IMP_MOVDOL) Then
                  .Cells(r_int_ConVer, 6) = 0
               Else
                  .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!IMP_MOVDOL)
                  r_dbl_Evalua(r_int_ConTem) = r_dbl_Evalua(r_int_ConTem) + Trim(g_rst_Princi!IMP_MOVDOL)
               End If
               .Cells(r_int_ConVer, 7) = 0
               
            ElseIf Trim(g_rst_Princi!FLAG_DEBHAB) = "H" Then
               .Cells(r_int_ConVer, 6) = 0
               If IsNull(g_rst_Princi!IMP_MOVDOL) Then
                  .Cells(r_int_ConVer, 7) = 0
               Else
                  .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!IMP_MOVDOL)
                  r_dbl_Evalua(r_int_ConTem + 1) = r_dbl_Evalua(r_int_ConTem + 1) + Trim(g_rst_Princi!IMP_MOVDOL)
               End If
            End If
            
         ElseIf Trim(cmb_TipMon.ListIndex) = 2 Then
            If Trim(g_rst_Princi!FLAG_DEBHAB) = "D" Then
               If IsNull(g_rst_Princi!IMP_MOVSOL) Then
                  .Cells(r_int_ConVer, 6) = 0
               Else
                  .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!IMP_MOVSOL)
                  r_dbl_Evalua(r_int_ConTem) = r_dbl_Evalua(r_int_ConTem) + Trim(g_rst_Princi!IMP_MOVSOL)
               End If
               .Cells(r_int_ConVer, 7) = 0
               
               If IsNull(g_rst_Princi!IMP_MOVDOL) Then
                  .Cells(r_int_ConVer, 8) = 0
               Else
                  .Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!IMP_MOVDOL)
                  r_dbl_Evalua(r_int_ConTem + 2) = r_dbl_Evalua(r_int_ConTem + 2) + Trim(g_rst_Princi!IMP_MOVDOL)
               End If
               .Cells(r_int_ConVer, 9) = 0
               
            ElseIf Trim(g_rst_Princi!FLAG_DEBHAB) = "H" Then
               .Cells(r_int_ConVer, 6) = 0
               If IsNull(g_rst_Princi!IMP_MOVSOL) Then
                  .Cells(r_int_ConVer, 7) = 0
               Else
                  .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!IMP_MOVSOL)
                  r_dbl_Evalua(r_int_ConTem + 1) = r_dbl_Evalua(r_int_ConTem + 1) + Trim(g_rst_Princi!IMP_MOVSOL)
               End If
               
               .Cells(r_int_ConVer, 8) = 0
               If IsNull(g_rst_Princi!IMP_MOVDOL) Then
                  .Cells(r_int_ConVer, 9) = 0
               Else
                  .Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!IMP_MOVDOL)
                  r_dbl_Evalua(r_int_ConTem + 3) = r_dbl_Evalua(r_int_ConTem + 3) + Trim(g_rst_Princi!IMP_MOVDOL)
               End If
               
            End If
            
         ElseIf Trim(cmb_TipMon.ListIndex) = 3 Then
            If Mid(g_rst_Princi!CNTA_CTBL, 3, 1) = 1 Then
               If Trim(g_rst_Princi!FLAG_DEBHAB) = "D" Then
                  If IsNull(g_rst_Princi!IMP_MOVSOL) Then
                     .Cells(r_int_ConVer, 6) = 0
                  Else
                     .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!IMP_MOVSOL)
                     r_dbl_Evalua(r_int_ConTem) = r_dbl_Evalua(r_int_ConTem) + Trim(g_rst_Princi!IMP_MOVSOL)
                  End If
                  .Cells(r_int_ConVer, 7) = 0
                  .Cells(r_int_ConVer, 8) = 0
                  .Cells(r_int_ConVer, 9) = 0
                  
               ElseIf Trim(g_rst_Princi!FLAG_DEBHAB) = "H" Then
                  .Cells(r_int_ConVer, 6) = 0
                  If IsNull(g_rst_Princi!IMP_MOVSOL) Then
                     .Cells(r_int_ConVer, 7) = 0
                  Else
                     .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!IMP_MOVSOL)
                     r_dbl_Evalua(r_int_ConTem + 1) = r_dbl_Evalua(r_int_ConTem + 1) + Trim(g_rst_Princi!IMP_MOVSOL)
                  End If
                  .Cells(r_int_ConVer, 8) = 0
                  .Cells(r_int_ConVer, 9) = 0
   
               End If
            ElseIf Mid(g_rst_Princi!CNTA_CTBL, 3, 1) = 2 Then
               If Trim(g_rst_Princi!FLAG_DEBHAB) = "D" Then
                  If IsNull(g_rst_Princi!IMP_MOVSOL) Then
                     .Cells(r_int_ConVer, 6) = 0
                  Else
                     .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!IMP_MOVSOL)
                     r_dbl_Evalua(r_int_ConTem) = r_dbl_Evalua(r_int_ConTem) + Trim(g_rst_Princi!IMP_MOVSOL)
                  End If
                  .Cells(r_int_ConVer, 7) = 0
                  
                  If IsNull(g_rst_Princi!IMP_MOVDOL) Then
                     .Cells(r_int_ConVer, 8) = 0
                  Else
                     .Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!IMP_MOVDOL)
                     r_dbl_Evalua(r_int_ConTem + 2) = r_dbl_Evalua(r_int_ConTem + 2) + Trim(g_rst_Princi!IMP_MOVDOL)
                  End If
                  .Cells(r_int_ConVer, 9) = 0
                  
               ElseIf Trim(g_rst_Princi!FLAG_DEBHAB) = "H" Then
                  .Cells(r_int_ConVer, 6) = 0
                  If IsNull(g_rst_Princi!IMP_MOVSOL) Then
                     .Cells(r_int_ConVer, 7) = 0
                  Else
                     .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!IMP_MOVSOL)
                     r_dbl_Evalua(r_int_ConTem + 1) = r_dbl_Evalua(r_int_ConTem + 1) + Trim(g_rst_Princi!IMP_MOVSOL)
                  End If
                  
                  .Cells(r_int_ConVer, 8) = 0
                  If IsNull(g_rst_Princi!IMP_MOVDOL) Then
                     .Cells(r_int_ConVer, 9) = 0
                  Else
                     .Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!IMP_MOVDOL)
                     r_dbl_Evalua(r_int_ConTem + 3) = r_dbl_Evalua(r_int_ConTem + 3) + Trim(g_rst_Princi!IMP_MOVDOL)
                  End If
   
               End If
            End If
         End If
                                                            
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      .Cells(r_int_ConVer, 5).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(r_int_ConVer, 5), .Cells(r_int_ConVer, r_int_Contad)).Font.Bold = True
      .Cells(r_int_ConVer, 5) = "TOTAL"
      .Cells(r_int_ConVer, 6) = r_dbl_Evalua(r_int_ConTem)
      .Cells(r_int_ConVer, 7) = r_dbl_Evalua(r_int_ConTem + 1)
      
      If Trim(cmb_TipMon.ListIndex) = 2 Or Trim(cmb_TipMon.ListIndex) = 3 Then
         .Cells(r_int_ConVer, 8) = r_dbl_Evalua(r_int_ConTem + 2)
         .Cells(r_int_ConVer, 9) = r_dbl_Evalua(r_int_ConTem + 3)
      End If
      
      For r_int_Contad = 0 To r_int_ConVer - 1 Step 4
         r_dbl_Evalua(30036) = r_dbl_Evalua(30036) + r_dbl_Evalua(r_int_Contad)
         r_dbl_Evalua(30037) = r_dbl_Evalua(30037) + r_dbl_Evalua(r_int_Contad + 1)
         r_dbl_Evalua(30038) = r_dbl_Evalua(30038) + r_dbl_Evalua(r_int_Contad + 2)
         r_dbl_Evalua(30039) = r_dbl_Evalua(30039) + r_dbl_Evalua(r_int_Contad + 3)
      Next
                  
      .Cells(r_int_ConVer + 2, 5).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(r_int_ConVer + 2, 5), .Cells(r_int_ConVer + 2, r_int_Contad)).Font.Bold = True
      .Cells(r_int_ConVer + 2, 5) = "TOTAL GENERAL"
      .Cells(r_int_ConVer + 2, 6) = r_dbl_Evalua(30036)
      .Cells(r_int_ConVer + 2, 7) = r_dbl_Evalua(30037)
      
      If Trim(cmb_TipMon.ListIndex) = 2 Or Trim(cmb_TipMon.ListIndex) = 3 Then
         .Cells(r_int_ConVer + 2, 8) = r_dbl_Evalua(30038)
         .Cells(r_int_ConVer + 2, 9) = r_dbl_Evalua(30039)
      End If
      
      .Range(.Cells(1, 1), .Cells(r_int_ConVer + 3, 99)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_ConVer + 3, 99)).Font.Size = 9
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

''********************************************************
Private Sub fs_GenExc_LMayor()
Dim r_obj_Excel      As Excel.Application
Dim r_int_nrofil     As Integer
Dim r_int_ConFil     As Integer
Dim r_str_CtaCtb     As String
Dim r_dbl_TotDeu     As Double
Dim r_dbl_TotAcr     As Double
    
   r_str_CtaCtb = ""
   r_int_ConFil = 2
   r_dbl_TotDeu = 0
   r_dbl_TotAcr = 0
    
   '**********************************
   '*********PROCEDURE****************
   '**********************************
   g_str_Parame = ""
   g_str_Parame = "usp_lm_libro_mayor ("
   g_str_Parame = g_str_Parame & "'" & IIf((cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) = 1, "sol", "dol") & "', "
   g_str_Parame = g_str_Parame & "'L',"
   g_str_Parame = g_str_Parame & "'" & ipp_FecIni.Text & "', "
   g_str_Parame = g_str_Parame & "'" & ipp_FecFin.Text & "', "
   g_str_Parame = g_str_Parame & "'L','LM')"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 2) Then
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
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT CNTA_CTBL, DESC_CNTA,TO_CHAR(FECHA_CNTBL,'DD/MM/YYYY') AS  FECHA_CNTBL, "
   g_str_Parame = g_str_Parame & "       NRO_LIBRO, TIPO_NOTA, NRO_ASIENTO, DESC_GLOSA, DEBE, HABER, ITEM "
   g_str_Parame = g_str_Parame & "  FROM LM_LIBRO_MAYOR "
   g_str_Parame = g_str_Parame & " ORDER BY CNTA_CTBL ASC,FECHA_CNTBL ASC, NRO_LIBRO ASC, NRO_ASIENTO ASC"
 
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
   r_obj_Excel.Sheets(1).Name = "Libro Mayor Dolares"
   If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 1 Then
       r_obj_Excel.Sheets(1).Name = "Libro Mayor Soles"
   End If
    
   With r_obj_Excel.Sheets(1)
      'PIE DE PAGINAS
      .PageSetup.CenterFooter = "&P de &N"
      
      'CENTRADO DE LA PAGINA
      .PageSetup.CenterHorizontally = True
      .PageSetup.Orientation = xlLandscape
      
      'AJUSTE DE ESCALA
      .PageSetup.Zoom = 70
      
      'IMPRIMIR TITULOS
      .PageSetup.PrintTitleRows = "$1:$10"
        
      'MARGENES
      .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
      .PageSetup.RightMargin = Application.CentimetersToPoints(1)
      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
      .Cells(1, 1) = "FORMATO 6.1: ""LIBRO MAYOR"""
      .Cells(2, 1) = "(MONEDA EXTRANJERA)"
      If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 1 Then
         .Cells(2, 1) = "(MONEDA NACIONAL)"
      End If
      
      .Cells(4, 1) = "PERIODO: "
      .Cells(4, 3) = "DEL " & Trim(ipp_FecIni) & " AL " & Trim(ipp_FecFin)
      .Cells(5, 1) = "RUC: "
      .Cells(5, 3) = "20511904162"
      .Cells(6, 1) = "DENOMINACIÓN O RAZÓN SOCIAL: "
      .Cells(6, 3) = "EDPYME MICASITA S.A."
       
      .Cells(9, 1) = "CUENTA"
      .Cells(9, 2) = "DESCRIPCIÓN"
      .Cells(9, 3) = "FECHA"
      .Cells(9, 4) = "LIBRO"
      .Cells(9, 5) = "NOTA"
      .Cells(9, 6) = "ASIENTO"
      .Cells(9, 7) = "DESCRIPCIÓN O GLOSA DE LA OPERACIÓN"
      .Cells(9, 8) = "SALDOS Y MOVIMIENTOS (US$.)"
      .Cells(10, 8) = "DEUDOR"
      .Cells(10, 9) = "ACREEDOR"
      
      If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) = 1 Then
         .Cells(9, 8) = "SALDOS Y MOVIMIENTOS (S/.)"
      End If
      
      .Range(.Cells(9, 8), .Cells(9, 9)).Merge
      .Range(.Cells(9, 1), .Cells(10, 9)).Font.Bold = True
      .Range(.Cells(9, 1), .Cells(10, 9)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").NumberFormat = "@"
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").NumberFormat = "###,###,##0.00"
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter:         .Columns("A").ColumnWidth = 14
      .Columns("B").ColumnWidth = 41
      .Columns("C").HorizontalAlignment = xlHAlignCenter:         .Columns("C").ColumnWidth = 13
      .Columns("D").HorizontalAlignment = xlHAlignCenter:         .Columns("D").ColumnWidth = 7
      .Columns("E").HorizontalAlignment = xlHAlignCenter:         .Columns("E").ColumnWidth = 7
      .Columns("F").HorizontalAlignment = xlHAlignCenter:         .Columns("F").ColumnWidth = 10
      .Columns("G").ColumnWidth = 56
      .Columns("H").ColumnWidth = 14
      .Columns("I").ColumnWidth = 14
      
      .Range(.Cells(1, 1), .Cells(6, 6)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(1, 1), .Cells(6, 6)).Font.Bold = True
      r_int_nrofil = 11
      
      .Range(.Cells(9, 1), .Cells(9, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(10, 8), .Cells(10, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(10, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(10, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(9, 2), .Cells(10, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(9, 3), .Cells(10, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(9, 4), .Cells(10, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(9, 5), .Cells(10, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(9, 6), .Cells(10, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(9, 7), .Cells(10, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(10, 8), .Cells(10, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(9, 9), .Cells(10, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
      
      .Range(.Cells(9, 1), .Cells(10, 9)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(9, 1), .Cells(10, 9)).WrapText = True
      .Range(.Cells(9, 1), .Cells(10, 9)).Font.Bold = True
      .Range(.Cells(9, 1), .Cells(10, 1)).Merge
      .Range(.Cells(9, 2), .Cells(10, 2)).Merge
      .Range(.Cells(9, 3), .Cells(10, 3)).Merge
      .Range(.Cells(9, 4), .Cells(10, 4)).Merge
      .Range(.Cells(9, 5), .Cells(10, 5)).Merge
      .Range(.Cells(9, 6), .Cells(10, 6)).Merge
      .Range(.Cells(9, 7), .Cells(10, 7)).Merge
      
      g_rst_Princi.MoveFirst
      r_str_CtaCtb = Trim(g_rst_Princi!CNTA_CTBL)
    
      Do While Not g_rst_Princi.EOF
         If Trim(g_rst_Princi!CNTA_CTBL) <> r_str_CtaCtb Then
            If r_int_ConFil = 54 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Van"
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(r_dbl_TotDeu)
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(r_dbl_TotAcr)
               r_int_nrofil = r_int_nrofil + 1
               
               'HorizontalPageBreaks
               r_obj_Excel.ActiveSheet.Rows(r_int_nrofil).PageBreak = xlPageBreakManual
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Vienen"
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(r_dbl_TotDeu)
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(r_dbl_TotAcr)
               r_int_nrofil = r_int_nrofil + 1
               r_int_ConFil = 0
            End If
            
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Saldo Final"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = r_dbl_TotDeu - r_dbl_TotAcr
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = 0
            If r_dbl_TotAcr > r_dbl_TotDeu Then
                r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = 0
                r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = r_dbl_TotAcr - r_dbl_TotDeu
            End If
            r_int_nrofil = r_int_nrofil + 1
            r_int_ConFil = r_int_ConFil + 1
            r_dbl_TotDeu = 0
            r_dbl_TotAcr = 0
         End If
 
         If r_int_ConFil = 54 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Van"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(r_dbl_TotDeu)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(r_dbl_TotAcr)
            r_int_nrofil = r_int_nrofil + 1
            
            'HorizontalPageBreaks
            r_obj_Excel.ActiveSheet.Rows(r_int_nrofil).PageBreak = xlPageBreakManual
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Vienen"
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(r_dbl_TotDeu)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(r_dbl_TotAcr)
            r_int_nrofil = r_int_nrofil + 1
            r_int_ConFil = 0
         End If
         
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 1) = Trim(g_rst_Princi!CNTA_CTBL)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 2) = Trim(g_rst_Princi!DESC_CNTA)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 3) = CDate(g_rst_Princi!FECHA_CNTBL)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 4) = Trim(g_rst_Princi!NRO_LIBRO)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 5) = Trim(g_rst_Princi!TIPO_NOTA)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 6) = Trim(g_rst_Princi!NRO_ASIENTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = Trim(g_rst_Princi!DESC_GLOSA)
         
         If CInt(g_rst_Princi!NRO_LIBRO) = 0 And CInt(g_rst_Princi!NRO_ASIENTO) = 0 And CInt(g_rst_Princi!Item) = 0 Then
            If CDbl(g_rst_Princi!DEBE) > CDbl(g_rst_Princi!HABER) Then
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(g_rst_Princi!DEBE) - CDbl(g_rst_Princi!HABER)
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = 0
               r_dbl_TotDeu = r_dbl_TotDeu + CDbl(g_rst_Princi!DEBE) - CDbl(g_rst_Princi!HABER)
            Else
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = 0
               r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(g_rst_Princi!HABER) - CDbl(g_rst_Princi!DEBE)
               r_dbl_TotAcr = r_dbl_TotAcr + CDbl(g_rst_Princi!HABER) - CDbl(g_rst_Princi!DEBE)
            End If
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(g_rst_Princi!DEBE)
            r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(g_rst_Princi!HABER)
            r_dbl_TotDeu = r_dbl_TotDeu + CDbl(g_rst_Princi!DEBE)
            r_dbl_TotAcr = r_dbl_TotAcr + CDbl(g_rst_Princi!HABER)
         End If
         
         r_int_nrofil = r_int_nrofil + 1
         r_int_ConFil = r_int_ConFil + 1
         r_str_CtaCtb = Trim(g_rst_Princi!CNTA_CTBL)
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      If r_int_ConFil = 54 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Van"
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(r_dbl_TotDeu)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(r_dbl_TotAcr)
         r_int_nrofil = r_int_nrofil + 1
         
         'HorizontalPageBreaks
         r_obj_Excel.ActiveSheet.Rows(r_int_nrofil).PageBreak = xlPageBreakManual
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Vienen"
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = CDbl(r_dbl_TotDeu)
         r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = CDbl(r_dbl_TotAcr)
         r_int_nrofil = r_int_nrofil + 1
         r_int_ConFil = 0
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 7) = "Saldo Final"
      r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = r_dbl_TotDeu - r_dbl_TotAcr
      r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = 0
      If r_dbl_TotAcr > r_dbl_TotDeu Then
          r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 8) = 0
          r_obj_Excel.ActiveSheet.Cells(r_int_nrofil, 9) = r_dbl_TotAcr - r_dbl_TotDeu
      End If
      
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 9)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_nrofil, 9)).Font.Size = 9
      .Rows("1:" & r_int_nrofil & "").RowHeight = 12
   End With
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipMon)
   End If
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Imprim)
   End If
End Sub
