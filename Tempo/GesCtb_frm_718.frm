VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RepSbs_20 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   10005
   ClientTop       =   3180
   ClientWidth     =   5325
   Icon            =   "GesCtb_frm_718.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5355
      _Version        =   65536
      _ExtentX        =   9446
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
         TabIndex        =   1
         Top             =   60
         Width           =   5265
         _Version        =   65536
         _ExtentX        =   9287
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
            TabIndex        =   2
            Top             =   30
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte N� 13"
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
            Top             =   270
            Width           =   4515
            _Version        =   65536
            _ExtentX        =   7964
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F0213-01 Control de Limites Globales e Individuales"
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
            Picture         =   "GesCtb_frm_718.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   780
         Width           =   5265
         _Version        =   65536
         _ExtentX        =   9287
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
            Picture         =   "GesCtb_frm_718.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_718.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_718.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4650
            Picture         =   "GesCtb_frm_718.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opci�n"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpDet 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_718.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   2610
            Top             =   90
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentaci�n Preliminar"
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
         TabIndex        =   10
         Top             =   1470
         Width           =   5265
         _Version        =   65536
         _ExtentX        =   9287
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
            TabIndex        =   11
            Top             =   90
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1530
            TabIndex        =   12
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
            Caption         =   "A�o:"
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_RepSbs_20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Dim r_dbl_Evalua(300)   As Double
   Dim r_str_Denomi(100)   As String
   Dim r_int_ConAux        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConTem        As Integer
   Dim r_int_TemAux        As Integer
   Dim r_int_PerMes        As String
   Dim r_int_PerAno        As String
   Dim r_dbl_MulUso        As Double
   Dim r_str_Cadena        As String
   Dim r_str_NomRes        As String
   Dim r_str_ParAux        As String
   Dim r_dbl_Volati        As Double
   Dim r_dbl_TipCam        As Double
   'Dim r_dbl_PatEfe        As Double
   Dim r_int_Cantid        As Integer
   Dim r_int_FlgRpr        As Integer
   
Private Sub cmd_Imprim_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   If ipp_PerAno.Text = "" Then
      MsgBox "Debe seleccionar el A�o.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("�Est� seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   
  Call fs_GenRpt
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   r_dbl_Volati = 0.01
      
   Screen.MousePointer = 0
End Sub
   
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Limpia()

   Dim r_int_PerMes  As Integer
   Dim r_int_PerAno  As Integer

   r_int_PerMes = Month(date)
   r_int_PerAno = Year(date)
   
   If Month(date) = 12 Then
      r_int_PerMes = 1
      r_int_PerAno = Year(date) - 1
   Else
      r_int_PerMes = Month(date) - 1
      r_int_PerAno = Year(date)
   End If
 
   Call gs_BuscarCombo_Item(cmb_PerMes, r_int_PerMes)
   ipp_PerAno.Text = Format(r_int_PerAno, "0000")
   
End Sub

Private Sub fs_Inicia()
         
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   
End Sub

Private Sub cmd_ExpArc_Click()

   Dim r_int_MsgBox As Integer
   
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   If ipp_PerAno.Text = "" Then
      MsgBox "Debe seleccionar el A�o.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   r_int_Cantid = modsec_gf_CanReg("HIS_CAMPAT", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
   If r_int_Cantid = 0 Then
      If MsgBox("�Est� seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      r_int_FlgRpr = 1
            
   Else
      r_int_MsgBox = MsgBox("�Desea reprocesar los datos?", vbQuestion + vbYesNoCancel + vbDefaultButton2, modgen_g_str_NomPlt)
      If r_int_MsgBox = vbNo Then
         r_int_FlgRpr = 0
         Call fs_GenArc
         
         Exit Sub
      ElseIf r_int_MsgBox = vbCancel Then
         Exit Sub
      ElseIf r_int_MsgBox = vbYes Then
         r_int_FlgRpr = 1
      End If
   
   End If
   
  Call fs_GenArc
End Sub

Private Sub cmd_ExpExc_Click()

   Dim r_int_MsgBox As Integer

   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   If ipp_PerAno.Text = "" Then
      MsgBox "Debe seleccionar el A�o.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   r_int_Cantid = modsec_gf_CanReg("HIS_CONLIM", CInt(ipp_PerAno.Text), CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)))
   
   If r_int_Cantid = 0 Then
      If MsgBox("�Est� seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      r_int_FlgRpr = 1
            
   Else
      r_int_MsgBox = MsgBox("�Desea reprocesar los datos?", vbQuestion + vbYesNoCancel + vbDefaultButton2, modgen_g_str_NomPlt)
      If r_int_MsgBox = vbNo Then
         r_int_FlgRpr = 0
         Call fs_GenExc
         Exit Sub
         
      ElseIf r_int_MsgBox = vbCancel Then
         Exit Sub
         
      ElseIf r_int_MsgBox = vbYes Then
         r_int_FlgRpr = 1
      End If
   
   End If
   
  Call fs_GenExc
  
End Sub

Private Sub fs_GenDat()

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   Erase r_str_Denomi
   Erase r_dbl_Evalua()
         
End Sub

Private Sub fs_GenExc()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_int_ConVar     As Integer
   Dim r_str_TipMon     As String
   
   Screen.MousePointer = 11
   
   If r_int_FlgRpr = 1 Then
      Call fs_GenDat
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      .Cells(1, 1) = "SUPERINTENDENCIA DE BANCA, SEGUROS Y AFP"
      .Cells(1, 4) = "REPORTE N�13"
                  
      .Cells(3, 1) = "CONTROL DE L�MITES GLOBALES E INDIVIDUALES APLICABLES A LAS EMPRESAS DEL SISTEMAS FINANCIERO"
      .Cells(4, 1) = "(Contenidas en la Ley N�26702 y normas reglamentarias emitidas por la SBS)"
      .Cells(5, 1) = "Al " & Left(modsec_gf_Fin_Del_Mes("01/" & r_int_PerMes & "/" & ipp_PerAno.Text), 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & ipp_PerAno.Text
      .Cells(7, 1) = "EMPRESA:Edpyme Micasita"
      
      .Cells(9, 1) = "I. L�mites Globales (1)"
      .Cells(10, 1) = "Aspecto"
      .Cells(10, 2) = "Base Legal"
      .Cells(10, 3) = "L�mites Legales"
      .Cells(10, 4) = "C�mputo"
      '.Cells(10, 5) = "Numerador"
      '.Cells(10, 6) = "Denominador"
      '.Cells(10, 7) = "Factor"
      '.Cells(10, 8) = "Contraparte"
      
      .Cells(11, 1) = "1. Ratio de Capital Global"
      .Cells(11, 2) = "Art�culo 199� y Vig�simo Cuarta Disposici�n Transitoria de la Ley Gral. Res.SBSN�2115-2009 Res. SBS N�4595-2009, N�4727-2009, N�4729-2009 y N�6328-2009"
      .Cells(11, 3) = "El patrimonio efectivo debe de ser mayor o igualal 10% de la suma de: Activos y Contingentes Ponderados por Riesgo de Cr�dito + 10 multiplicado por Requerimiento " & _
                     "de Patrimonio Efectivo por Riesgo de Mercado + 10 multiplicado por Requerimiento de Patrimonio Efectivo por Riesgo Operacional. (2)"
      .Cells(11, 4) = Format(r_dbl_Evalua(2), "##0.00") & "% de los Activos y Contingentes Ponderadospor Riesgos Totales.%"
      
      .Cells(12, 1) = "2.1. Componentes de Patrimonio B�sico"
      .Cells(12, 2) = "Art�culo 184�, �ltimo p�rrafo literalA Ley Gral. Res. SBS N�4595-2009,  N�4727-2009 y N�4729-2009"
      .Cells(12, 3) = "Los elementos del numeral 3 del literal A del Art.184� de la Ley Gral. s�lo deber�n computarse hasta el 17.65% del monto correspondiente a los componentes considerados " & _
                     "en los numerales 1, 2, 4 y 5 del mismo literal."
      .Cells(12, 4) = Format(r_dbl_Evalua(5), "##0.00") & "% del monto correspondiente a los componentes considerados en los numerales 1, 2, 4 y 5 del literal A del art�culo 184�."
      
      .Cells(13, 1) = "2.2. Patrimonio Suplementario"
      .Cells(13, 2) = "Art�culos 184� y N�185� Ley General Res. SBS N�4595-2009, N�4727-2209 y N�4529-2009"
      .Cells(13, 3) = "El patrimonio suplementario debe de ser menor o igual al 100% del Patrimonio B�sico."
      .Cells(13, 4) = Format(r_dbl_Evalua(8), "##0.00") & "% del patrimonio b�sico de la empresa"
      
      .Cells(14, 1) = "2.3..... Deuda Subordinada Redimible del Patrimonio de Nivel 2"
      .Cells(14, 2) = "Art�culos 184� y N�185� Ley General Res. SBS N�4595-2009, N�4727-2209 y N�4529-2009"
      .Cells(14, 3) = "La deuda subordinada redimible del patrimonio de nivel 2 debe de ser menor o igual al 50% del monto correspondiente a los componentes del patrimonio b�sico considerados " & _
                     "en los numerales 1, 2, 3,4  y 5 del literal A del art�culo 184�."
      .Cells(14, 4) = Format(r_dbl_Evalua(11), "##0.00") & "% de los componentes del patrimonio b�sico de la empresa, considerados en los numerales 1, 2, 3, 4 y 5 del literal A del art�culo 184�."
      
      .Cells(15, 1) = "2.4. Patrimonio Nivel 3"
      .Cells(15, 2) = "Art�culos 184� y N�185� Ley General Res. SBS N�4595-2009, N�4727-2209 y N�4529-2009"
      .Cells(15, 3) = "El patrimonio de nivel 3 debe de ser menor o igual al 250% del monto correspondiente a los componentes del patrimonio b�sico considerados en los numerales 1, 2, 3, 4 y 5 " & _
                     " del literal A del art�culo 184� asignados a cubrir riesgos de mercado."
      .Cells(15, 4) = Format(r_dbl_Evalua(14), "##0.00") & "% de los componentes del patrim�nio b�sico de la empresa, considerados en los numerales 1, 2, 3, 4 y 5 del literal A del art�culo 184� asignado a cubrir riesgos de mercado."
      
      .Cells(16, 1) = "3.1. Riesgos de Mercado - Operaciones con productos financieros derivados"
      .Cells(16, 2) = "Art�culo 200�, num. 3 Ley Gral. Resoluci�n SBS N�1737-2006"
      .Cells(16, 3) = "Hasta el 10% del patrimonio efectivo."
      .Cells(16, 4) = Format(r_dbl_Evalua(17), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(17, 1) = "3.2. Riesgo de Mercado - Tenencias en acciones cotizadas en mecanismos centralizados de negociaci�n, as� como los certificados de participaci�n en fondos mutuos y los certificados " & _
                     " de participaci�n en fondos de inversi�n."
      .Cells(17, 2) = "Art�culo 200�, num. 4 Ley Gral. Resoluci�n SBS N�1067-2005"
      .Cells(17, 3) = "Hasta el 40% del patrimonio efectivo."
      .Cells(17, 4) = Format(r_dbl_Evalua(20), "##0.00") & "% del patrimonioefectivo de la empresa."
      
      .Cells(18, 1) = "4. Requerimiento m�nimo de liquidez en moneda nacional"
      .Cells(18, 2) = "Resoluci�n SBS N�472-2201"
      .Cells(18, 3) = "Activos L�quidos: Mayor o igual al 8% de los pasivos de corto plazo en M.N.(Promedio mensual calendario de saldos diarios)"
      .Cells(18, 4) = Format(r_dbl_Evalua(23), "##0.00") & "% de los pasivos de corto plazo en M.N."
      
      .Cells(19, 1) = "5. Requerimiento m�nimo de liquidez"
      .Cells(19, 2) = "Resoluci�n SBS N�472-2201"
      .Cells(19, 3) = "Activos L�quidos: Mayor o igual al 20% de los pasivosde corto plazo en M.E. (Promedio mensual calendrio de saldos diarios)"
      .Cells(19, 4) = Format(r_dbl_Evalua(26), "##0.00") & "% de los pasivos de corto plazo en M.E."
      
      .Cells(20, 1) = "6. Inversiones en capital social de subsidiarias"
      .Cells(20, 2) = "Art�culo 36�, num. 1 Ley Gral."
      .Cells(20, 3) = "Hasta el 40% del patrimonio contable de la empresa."
      .Cells(20, 4) = Format(r_dbl_Evalua(29), "##0.00") & "% del patrimonio contable de la empresa"

      .Cells(21, 1) = "7. Adquisi�n de facturas mediante factoring"
      .Cells(21, 2) = "Art�culo 200�, num. 1 Ley Gral. Resoluci�n SBS N�1021-98"
      .Cells(21, 3) = "Hasta el 15% del patrimonio efectivo."
      .Cells(21, 4) = Format(r_dbl_Evalua(32), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(22, 1) = "8. Tenencia de oro"
      .Cells(22, 2) = "Art�culo 200�, num. 2 Ley Gral."
      .Cells(22, 3) = "Hasta el 15% del patrimonio efectivo."
      .Cells(22, 4) = Format(r_dbl_Evalua(35), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(23, 1) = "9. Inversi�n en letras hipotecarias de propia emisi�n"
      .Cells(23, 2) = "Circular SBS N�B-1959-94 y similares"
      .Cells(23, 3) = "Hasta el 5% del patrimonio efectivo."
      .Cells(23, 4) = Format(r_dbl_Evalua(38), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(24, 2) = "Circular SBS N�B-1959-94 y similares"
      .Cells(24, 3) = "Excepcionalmente hasta el 10% del patrimonio efectivo previa autorizaci�n de esta Superintendencia, sin exceder el l�mite se�alado en el numeral 7.1 de " & _
                     "Circular SBS N�B-1959-94."
      .Cells(24, 4) = Format(r_dbl_Evalua(41), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(25, 1) = "10. Inversi�n en muebles e inmuebles"
      .Cells(25, 2) = "Art�culo 200�, num. 6 Ley Gral. Resoluci�n SBS N�831-98"
      .Cells(25, 3) = "Hasta el 75% del patrimonio efectivo."
      .Cells(25, 4) = Format(r_dbl_Evalua(44), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(26, 1) = "10.1. Inversi�n en inmuebles"
      .Cells(26, 2) = "Resoluci�n SBS N�831-98"
      .Cells(26, 3) = "Subl�mite 40% del patrimonio efectivo para inversi�n en muebles. (3)"
      .Cells(26, 4) = Format(r_dbl_Evalua(47), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(27, 1) = "10.2. Inversi�n en muebles"
      .Cells(27, 2) = "Resoluci�n SBS N�831-98"
      .Cells(27, 3) = "Subl�mite 40% del patrimonio efectivo para inversi�n en inmuebles. (3)"
      .Cells(27, 4) = Format(r_dbl_Evalua(50), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(28, 1) = "11.1. L�mite la posici�n global de sobreventa de moneda extranjera"
      .Cells(28, 2) = "Resoluci�n SBS N�1455-2003"
      .Cells(28, 3) = "Hasta el 10% del patrimonio efectivo."
      .Cells(28, 4) = Format(r_dbl_Evalua(53), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(29, 1) = "11.2. L�mite la posici�n global de sobrecompra de moneda extranjera"
      .Cells(29, 2) = "Resoluci�n SBS N�1455-2003"
      .Cells(29, 3) = "Hasta el 100% del patrimonio efectivo."
      .Cells(29, 4) = Format(r_dbl_Evalua(56), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(30, 1) = "12. Cr�ditos a directos y trabajadores de la empresa"
      .Cells(30, 2) = "Total cr�ditos a directores y trabajadores de la empresa Art�culo 201� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(30, 3) = "Hasta el 7% del patrimonio efectivo."
      .Cells(30, 4) = Format(r_dbl_Evalua(59), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(31, 1) = "12. Cr�ditos a directos y trabajadores de la empresa"
      .Cells(31, 2) = "Total cr�ditos a directores y trabajadores de la empresa Art�culo 201� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(31, 3) = "Hasta el 7% del patrimonio efectivo."
      .Cells(31, 4) = Format(r_dbl_Evalua(62), "##0.00") & "% del patrimonio efectivo de la empresa."
   
      .Cells(32, 1) = "13. Financiamiento a personas vinculadas a la empresa"
      .Cells(32, 2) = "Art�culo 202� Ley Gral. Res. SBS 445-2000 y N�472-2006 Circular N�B-2148-2005"
      .Cells(32, 3) = "Hasta el 30% del patrimonio efectivo."
      .Cells(32, 4) = Format(r_dbl_Evalua(65), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(33, 1) = "14. Total de financiamientos a soberanos"
      .Cells(33, 2) = "Art�culo 203� de la Ley Gral. Art�culo 212� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(33, 3) = ""
      .Cells(33, 4) = Format(r_dbl_Evalua(68), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(34, 1) = "15. Total de financiamientos a entidades que realizan actividad empresarial del Estado, sin considerar aquellas empresas cuya autonom�a econ�mica y administrativa ha sido declarada por ley."
      .Cells(34, 2) = "Art�culo 203� de la Ley Gral. Art�culo 206� de la Ley Gral. Art�culo 207� de la Ley Gral. Art�culo 208� de la Ley Gral. Art�culo 209� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(34, 3) = "Hasta el 10% del patrimonio efectivo Hasta el 15% del patrimonio efectivo Hasta el 20% del patrimonio efectivo Hasta el 30% del patrimonio efectivo (Sujeto al tipo de garant�a)"
      .Cells(34, 4) = Format(r_dbl_Evalua(71), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(35, 1) = "16. Tota de financiamientos otorgados a otras entidades, organismos y dependencias que directa o indirectamente sean considerados o formen parte del Estado Peruano. No se incluyen los financiamientos " & _
                     "se�alados en los numerales 14 y 15 anteriores, ni los otorgados a los gobiernos locales o regionales, ni a COFIDE, AGROBANCO, Fondo MIVIVIENDA, Banco de la Naci�n y Cajas Municipales."
      .Cells(35, 2) = "Art�culo 203� de la Ley Gral. Art�culo 206� de la Ley Gral. Art�culo 207� de la Ley Gral. Art�culo 208� de la Ley Gral. Art�culo 209� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(35, 3) = "Hasta el 10% del patrimonio efectivo Hasta el 15% del patrimonio efectivo Hasta el 20% del patrimonio efectivo Hasta el 30% del patrimonio efectivo (Sujeto al tipo de garant�a)"
      .Cells(35, 4) = Format(r_dbl_Evalua(74), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      For r_int_Contad = 0 To 24 Step 1
         .Range(.Cells(r_int_Contad + 11, 1), .Cells(r_int_Contad + 11, 8)).WrapText = True
      Next
      
      .Range(.Cells(1, 8), .Cells(1, 4)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(1, 1), .Cells(3, 8)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(5, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 1)).Font.Bold = True
                  
      .Range(.Cells(1, 1), .Cells(1, 2)).Merge
      .Range(.Cells(3, 1), .Cells(3, 4)).Merge
      .Range(.Cells(4, 1), .Cells(4, 4)).Merge
      .Range(.Cells(5, 1), .Cells(5, 4)).Merge
      .Range(.Cells(23, 1), .Cells(24, 1)).Merge
      .Range(.Cells(9, 1), .Cells(9, 4)).Merge
                  
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(10, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(9, 1), .Cells(35, 8)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                 
      .Range(.Cells(9, 1), .Cells(10, 8)).Font.Bold = True
      
      
      .Cells(37, 1) = "II. L�mites Individuales(1)"
      .Cells(38, 1) = "Aspecto (4)"
      .Cells(38, 2) = "Base Legal"
      .Cells(38, 3) = "L�mites Legales"
      .Cells(38, 4) = "C�mputo (6)"
      '.Cells(38, 5) = "Numerador"
      '.Cells(38, 6) = "Denominador"
      '.Cells(38, 7) = "Factor"
      '.Cells(38, 8) = "Contraparte"
      
      .Cells(39, 1) = "1. Total de financiamiento a un deudor o grupo de deudores que representa riesgo �nico (5)."
      .Cells(39, 2) = "Art. 203�; 204�; 205�; 206�; 207�; 208�; 209�; 210�; 211� y 212� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(39, 3) = "Hasta el 30% del patrimonio efectivo debiendose adem�s tener encuenta los subl�mites contemplados en el numeral 4 de las Circular N�B-2148-2005. Hasta el 50%, si el exceso se encuentra representado " & _
                     "por cartas de cr�dito de empresas del sistema financiero del exterior de conformidad con elnumeral 4 del Art. 205� de la Ley Gral."
      .Cells(39, 4) = "1)" & Format(r_dbl_Evalua(77), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo de deudores que representa riesgo �nico, con el mayor financiamiento."
      .Cells(40, 4) = "2)" & Format(r_dbl_Evalua(80), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo de deudores que representa riesgo �nico, con el segundo mayor financiamiento."
      .Cells(41, 4) = "3)" & Format(r_dbl_Evalua(83), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo de deudores que representa riesgo �nico, con el tercer mayor financiamiento."
      
      .Cells(42, 1) = "2. Financiamiento directo e indirecto a empresas del sistema  financiero establecidas en el pa�s, dep�sitos constituidos en ellas, avales, fianzas y otras garant�as que se haya recibido de dichas empresas."
      .Cells(42, 2) = "Art�culo 204�; p�rrafo 1 Ley Gral. Art�culo 203� Ley Gral. Art. 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(42, 3) = "Hasta el 30% del patrimonio efectivo."
      .Cells(42, 4) = "1)" & Format(r_dbl_Evalua(86), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(43, 4) = "2)" & Format(r_dbl_Evalua(89), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(44, 4) = "3)" & Format(r_dbl_Evalua(92), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(45, 1) = "3. Financiamiento directo e indirecto a empresas bancarias o financieras del exterior, los dep�sitos constituidos en ellas y los avales avales, fianzas y otras garant�as que se haya recibido de dichas instituciones."
      .Cells(45, 2) = "Art�culo 204�; p�rrafo 1 Ley Gral. Art�culo 203� Ley Gral. Art. 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(45, 3) = "Hasta el 5% del patrimonio efectivo, en caso de empresas no sujetas a supervisi�n por organismos similares a la sbs. Hasta el 50% del patrimonio efectivo, siempre que el exceso se encuentre representado por cartas de cr�dito, incluyendo la modalidad de stand by letter of credit."
      .Cells(45, 4) = "1)" & Format(r_dbl_Evalua(95), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(46, 4) = "2)" & Format(r_dbl_Evalua(98), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(47, 4) = "3)" & Format(r_dbl_Evalua(101), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(48, 2) = "Art�culo 205�; numeral 2 y 4, 203� y 212� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(48, 3) = "Se puede exceder el l�mite anterior de 5% hasta el 10% del patrimonio efectivo, en caso de empresas sujetas a supervisi�n por organismos similares a ls SBS, y no son bancos de 1ra categor�a. Hasta el 50% del patrimonio efectivo, siempre que el exceso se encuentre representado por cartas de cr�dito, incluyendo la modalidad de stand by letter of credit."
      .Cells(48, 4) = "1)" & Format(r_dbl_Evalua(104), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(49, 4) = "2)" & Format(r_dbl_Evalua(107), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(50, 4) = "3)" & Format(r_dbl_Evalua(110), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(51, 2) = "Art�culo 205�; numeral 3 y 4, 203� y 212� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(51, 3) = "Se puede exceder el l�mite anterior de 5% y 10% anteriores hasta el 30% del patrimonio efectivo en caso de bancos de 1ra categor�a. Hasta el 50% del patrimonio efectivo, siempre que el exceso se encuentre representado por cartas de cr�dito, incluyendo la modalidad de stand by letter of credit."
      .Cells(51, 4) = "1)" & Format(r_dbl_Evalua(113), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(52, 4) = "2)" & Format(r_dbl_Evalua(116), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(53, 4) = "3)" & Format(r_dbl_Evalua(119), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(54, 1) = "4. Financiamiento mediante cr�ditos, inversiones y contingentes a favor de una misma persona natural jur�dica directa o indirectamente (para residentes en el pa�s con excepci�n de empresas del sistema financiero y de las entidades consideradas en los numerales 14, 15 y 16 de la selecci�n L�mites Globales)"
      .Cells(54, 2) = "Art�culo 206� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(54, 3) = "Hasta el 10% del patrimonio efectivo."
      .Cells(54, 4) = "1)" & Format(r_dbl_Evalua(122), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(55, 4) = "2)" & Format(r_dbl_Evalua(125), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(56, 4) = "3)" & Format(r_dbl_Evalua(128), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(57, 2) = "Art�culo 207� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(57, 3) = "Se puede exceder el l�mite contemplado en el primer p�rrafo del art�culo 206� de la Ley General hasta el 15% del patrimonio efectivo (sujeto al tipo de garant�a)"
      .Cells(57, 4) = "1)" & Format(r_dbl_Evalua(131), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(58, 4) = "2)" & Format(r_dbl_Evalua(134), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(59, 4) = "3)" & Format(r_dbl_Evalua(137), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(60, 2) = "Art�culo 208� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(60, 3) = "Se puede exceder los l�mite contemplados en el primer p�rrafo del Art. 206� y en el Art. 207� de la Ley General hasta el 20% del patrimonio efectivo (sujeto al tipo de garant�a)"
      .Cells(60, 4) = "1)" & Format(r_dbl_Evalua(140), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(61, 4) = "2)" & Format(r_dbl_Evalua(143), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(62, 4) = "3)" & Format(r_dbl_Evalua(146), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(63, 2) = "Art�culo 209� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(63, 3) = "Se puede exceder los l�mite contemplados en el primer p�rrafo del Art. 206�; en el Art. 207� y en el Art. 208� de la Ley General hasta el 30% del patrimonio efectivo (sujeto al tipo de garant�a)"
      .Cells(63, 4) = "1)" & Format(r_dbl_Evalua(149), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(64, 4) = "2)" & Format(r_dbl_Evalua(152), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(65, 4) = "3)" & Format(r_dbl_Evalua(155), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."

      .Cells(66, 1) = "5. Cr�ditos, contingentes, inversiones y arrendatarios financieros otorgados a una misma persona natural o jur�dica residentes en el exterior (conexcepcion de empresas del sistema financiero)"
      .Cells(66, 2) = "Art�culo 211� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(66, 3) = "Hasta el 5% del patrimonio efectivo."
      .Cells(66, 4) = "1)" & Format(r_dbl_Evalua(158), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(67, 4) = "2)" & Format(r_dbl_Evalua(161), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(68, 4) = "3)" & Format(r_dbl_Evalua(164), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."

      .Cells(69, 2) = "Art�culo 211� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(69, 3) = "Se puede exceder los l�mites anteriores de 5% y 10% hasta el 30% del patrimonio efectivo (Sujeto al tipo de garant�a)."
      .Cells(69, 4) = "1)" & Format(r_dbl_Evalua(167), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(70, 4) = "2)" & Format(r_dbl_Evalua(170), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(71, 4) = "3)" & Format(r_dbl_Evalua(173), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(72, 1) = "6. Fianzas otorgadas para garantizar la suscripci�n de contratos derivados de procesos de licitaci�n p�blica."
      .Cells(72, 2) = "Art�culo 206� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(72, 3) = "Hasta el 30% del patrimonio efectivo."
      .Cells(72, 4) = "1)" & Format(r_dbl_Evalua(176), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(73, 4) = "2)" & Format(r_dbl_Evalua(179), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(74, 4) = "3)" & Format(r_dbl_Evalua(182), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(75, 1) = "7. Inversiones en letras hipotecarias de un mismo emisor."
      .Cells(75, 2) = "Circular N�B-1959-94 y similares"
      .Cells(75, 3) = "Hasta el 10% del patrimonio efectivo."
      .Cells(75, 4) = "1)" & Format(r_dbl_Evalua(185), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya adquirido del emisor con el mayor financiamiento."
      .Cells(76, 4) = "2)" & Format(r_dbl_Evalua(188), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya adquirido del emisor con el segundo mayor financiamiento."
      .Cells(77, 4) = "3)" & Format(r_dbl_Evalua(191), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya adquirido del emisor con el tercer mayor financiamiento."
      
      .Cells(78, 1) = "8. Cr�dito a directores y trabajadores de la empresa."
      .Cells(78, 2) = "Art�culo 201� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(78, 3) = "Hasta el 5% de lo se�alado en el numeral 12 de la secci�n L�mites Globales (es decir 0.35% del patrimonio efectivo)."
      .Cells(78, 4) = "1)" & Format(r_dbl_Evalua(194), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(79, 4) = "2)" & Format(r_dbl_Evalua(197), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(80, 4) = "3)" & Format(r_dbl_Evalua(200), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(81, 1) = "9. Warrants recibidos en garant�a."
      .Cells(81, 2) = "Art�culo 204�; p�rrafo 2 Ley Gral. Circular N�B-2148-2005"
      .Cells(81, 3) = "Hasta el 60% del patrimonio efectivo."
      .Cells(81, 4) = "1)" & Format(r_dbl_Evalua(203), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del AGD con el mayor importe."
      .Cells(82, 4) = "2)" & Format(r_dbl_Evalua(206), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del AGD con el segundo mayor importe."
      .Cells(83, 4) = "3)" & Format(r_dbl_Evalua(209), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del AGD con el tercer mayor importe."
      
      .Cells(84, 1) = "10. Coberturas que otorguen un patrimonio aut�nomo de seguro de cr�dito o un fondo de garant�a creado por el Estado a favor de una misma empresa."
      .Cells(84, 2) = "Art�culo 204�; p�rrafo 3 Ley Gral. Art�culo 212� Ley Gral."
      .Cells(84, 3) = ""
      .Cells(84, 4) = "1)" & Format(r_dbl_Evalua(212), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del patrimonio aut�nomo o fondo con la mayor cobertura otorgada."
      .Cells(85, 4) = "2)" & Format(r_dbl_Evalua(215), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del patrimonio aut�nomo o fondo con la segundo mayor cobertura otorgada."
      .Cells(86, 4) = "3)" & Format(r_dbl_Evalua(218), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del patrimonio aut�nomo o fondo con la tercera mayor cobertura otorgada."
      
      .Cells(87, 1) = "11. Inversi�n en el capital social de una subsidiaria."
      .Cells(87, 2) = "Art�culo 36�; num. 2 Ley Gral."
      .Cells(87, 3) = "Mayor o igual a 3/5 partes del capital social de la subsidiaria."
      .Cells(87, 4) = Format(r_dbl_Evalua(221), "##0.00") & " partes del capital social de la subsidiaria."
      .Cells(88, 4) = Format(r_dbl_Evalua(224), "##0.00") & "Calcular el l�mite para cada subsidiaria."
      
      For r_int_Contad = 39 To 88 Step 1
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 8)).WrapText = True
      Next
      
      For r_int_Contad = 39 To 87 Step 3
         If r_int_Contad = 45 Then
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 8, 1)).Merge
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 8, 1)).WrapText = True
            r_int_Contad = r_int_Contad + 6
         ElseIf r_int_Contad = 54 Then
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 11, 1)).Merge
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 11, 1)).WrapText = True
            r_int_Contad = r_int_Contad + 9
         ElseIf r_int_Contad = 66 Then
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 5, 1)).Merge
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 5, 1)).WrapText = True
            r_int_Contad = r_int_Contad + 3
         ElseIf r_int_Contad <> 87 Then
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 2, 1)).Merge
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 2, 1)).WrapText = True
         Else
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 1, 1)).Merge
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 1, 1)).WrapText = True
         End If
      Next
            
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(38, 1), .Cells(38, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(37, 1), .Cells(95, 8)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                 
      .Range(.Cells(37, 1), .Cells(38, 8)).Font.Bold = True
      
      For r_int_Contad = 40 To 88 Step 3
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 8)).Borders(xlEdgeTop).LineStyle = xlNone
         If r_int_Contad <> 88 Then
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 8)).Borders(xlEdgeBottom).LineStyle = xlNone
         End If
      Next
      
      .Range(.Cells(90, 1), .Cells(90, 4)).Merge
      .Range(.Cells(91, 1), .Cells(91, 4)).Merge
      .Range(.Cells(92, 1), .Cells(92, 4)).Merge
      .Range(.Cells(93, 1), .Cells(93, 4)).Merge
      .Range(.Cells(94, 1), .Cells(94, 4)).Merge
      .Range(.Cells(95, 1), .Cells(95, 4)).Merge
      
      .Range(.Cells(90, 1), .Cells(90, 1)).RowHeight = 30
      '.Range(.Cells(92, 1), .Cells(92, 1)).RowHeight = 30
      .Range(.Cells(93, 1), .Cells(93, 1)).RowHeight = 30
      .Range(.Cells(94, 1), .Cells(94, 1)).RowHeight = 45
      
      .Range(.Cells(90, 1), .Cells(90, 1)).WrapText = True
      .Range(.Cells(92, 1), .Cells(92, 1)).WrapText = True
      .Range(.Cells(93, 1), .Cells(93, 1)).WrapText = True
      .Range(.Cells(94, 1), .Cells(94, 1)).WrapText = True
            
      .Cells(90, 1) = "(1) El patrimonio efectivo que deber� emplearse para el c�mputo de los l�mites ser� el �ltimo remitido por la empresa y validado por esta Superintendencia, salvo para el caso " & _
                     "del ratio de capital global y los subl�mites de los componentes del patrimonio efectivo."
      .Cells(91, 1) = "(2) Se deber� considerar el cronograma de adecuaci�n se�alado en la Vig�simo Cuarta Disposici�n Transitoria de la Ley General."
      .Cells(92, 1) = "(3) Se deber� considerar los saldos netos de depresaci�n y p�rdida por deterioro acumuladas. No se incluyen los bienes dados en arrendamiento financiero, bienes recuperados, " & _
                     "bienes adjudicados y recibidos en pago."
      .Cells(93, 1) = "(4) Para cada l�mite y/o subl�mite individual indicar tres (3) mayores exposiciones. Por ejemplo: para la medici�n del l�mite individual a que se refiere el art�culo 206�, " & _
                     "la empresa deber� reportar los tres (3) mayores financiamientos otorgados sin garant�as al deudor o grupo de deudores domiciliados en el pa�s qu representan riesgo �nico."
      .Cells(94, 1) = "(5) Las empresasdeber�naplicarlos l�mites de concentraci�n a que se refiere en los art�culos 204� al 211� de la Ley General considerando el criterio de riesgo �nico de conformidad " & _
                     "de conformidad con los dispuesto en el art�culo 203� de la Ley General y en el cap�tulo II de las Normas Especiales sobre Vinculaci�n y Grupo Econ�mico, de tal forma que un grupo " & _
                     "de contrapartes relacionadas que representen riesgo �nico no podr� exceder del treinta por ciento (30%) del patrimonio efectivo de la empresa, de conformidad con lo establecido en " & _
                     "el numeral 4 de la Circular B-2148-2005 y modificatorias."
      .Cells(95, 1) = "(6) Luego de reportar los indicadores de exposici�n se deber� incluir el nombre de la contraparte (persona o grupo, AGD,subsidiaria, patrimonio aut�nomo o fondo de garant�a)."
      
      
      .Cells(100, 2) = "_________________________"
      .Cells(101, 2) = "Gerente General"
      
      .Cells(100, 3) = "______________________________"
      .Cells(101, 3) = "Gerente de Unidad de Riesgos"
      
      .Range(.Cells(100, 2), .Cells(101, 3)).HorizontalAlignment = xlHAlignCenter
                 
                                              
      .Range(.Cells(1, 1), .Cells(110, 110)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(110, 110)).Font.Size = 9
            
      .Columns("A").ColumnWidth = 49
      .Columns("B").ColumnWidth = 49
      .Columns("C").ColumnWidth = 49
      .Columns("D").ColumnWidth = 49
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 15
      .Columns("G").ColumnWidth = 15
      .Columns("H").ColumnWidth = 15
                  
      .Columns("E:H").NumberFormat = "###,###,##0.00"
                         
   End With

   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
   
End Sub

Private Sub fs_GenArc()
  
   Dim r_int_NumRes     As Integer
   Dim r_int_CodEmp     As Integer
     
   Dim r_str_Cadena     As String
   Dim r_str_NomRes     As String
   Dim r_str_FecRpt     As String
   
   Dim r_dbl_MulUso     As Double
   
   If r_int_FlgRpr = 1 Then
      Call fs_GenDat
      Call fs_GeneDB
   ElseIf r_int_FlgRpr = 0 Then
      Call fs_GenDat_DB
   End If
      
   r_str_FecRpt = "01/" & Format(r_int_PerMes, "00") & "/" & r_int_PerAno
      
   r_str_NomRes = "C:\01" & Right(r_int_PerAno, 2) & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & ".213"
   
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
   
   Print #r_int_NumRes, Format(213, "0000") & Format(1, "00") & Format(r_int_CodEmp, "00000") & r_int_PerAno & Format(r_int_PerMes, "00") & Format(modsec_gf_Fin_Del_Mes(r_str_FecRpt), "dd") & Format(12, "000")
   
   r_int_ConAux = 0
   r_int_TemAux = 0
   
   For r_int_Contad = 100 To 7600 Step 100
      r_str_Cadena = ""
      
      If r_int_Contad <> 100 And r_int_Contad <> 2400 And r_int_Contad <> 7600 Then
   
         For r_int_ConTem = 0 To 2 Step 1
            If r_int_ConTem = 2 Then
               r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 10)
            Else
               r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_dbl_Evalua(r_int_ConAux), "########0.00"), 1, "0", 18)
            End If
            r_int_ConAux = r_int_ConAux + 1
         Next
         
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(r_str_Denomi(r_int_TemAux), "########0.00"), 2, " ", 40)
         r_int_TemAux = r_int_TemAux + 1
      
      Else
         For r_int_ConTem = 0 To 2 Step 1
            If r_int_ConTem = 2 Then
               r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 10)
            Else
               r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 18)
            End If

         Next
         
         r_str_Cadena = r_str_Cadena & gs_modsec_Genera(Format(" ", "########0.00"), 2, " ", 40)
            
      End If
      
      Print #r_int_NumRes, Format(r_int_Contad, "000000") & r_str_Cadena
      
      
      If r_int_Contad = 700 Or r_int_Contad = 2300 Or r_int_Contad = 750 Or r_int_Contad = 2350 Then
         r_int_Contad = r_int_Contad - 50
      
      End If
      
   Next
          
   'Cerrando Archivo Resumen
   Close #r_int_NumRes
   
   Screen.MousePointer = 0
   
   MsgBox "Archivo creado.", vbInformation, modgen_g_str_NomPlt

End Sub

Private Sub fs_GeneDB()

   If (r_int_PerMes <> IIf(Format(Now, "MM") - 1 = 0, 12, Format(Now, "MM") - 1)) Or (r_int_PerAno <> IIf(Format(Now, "MM") - 1 = 0, Format(Now, "YYYY") - 1, Format(Now, "YYYY"))) Then
      MsgBox "Periodo cerrado, no se guardar�n los datos.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   g_str_Parame = "DELETE FROM HIS_CONLIM WHERE "
   g_str_Parame = g_str_Parame & "CONLIM_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "CONLIM_PERANO = " & r_int_PerAno & " AND "
   g_str_Parame = g_str_Parame & "CONLIM_USUCRE = '" & modgen_g_str_CodUsu & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_ConAux = 0
   
   For r_int_Contad = 0 To 72 Step 1
            
      r_str_Cadena = "USP_HIS_CONLIM ("
      r_str_Cadena = r_str_Cadena & "'CTB_REPSBS_07', "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & 0 & ", "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
      r_str_Cadena = r_str_Cadena & CInt(r_int_PerMes) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_PerAno) & ", "
      r_str_Cadena = r_str_Cadena & CInt(r_int_Contad + 1) & ", "
      
      For r_int_ConTem = 0 To 2 Step 1
         r_str_Cadena = r_str_Cadena & ", " & r_dbl_Evalua(r_int_ConAux)
         r_int_ConAux = r_int_ConAux + 1
      Next
      
      r_str_Cadena = r_str_Cadena & ", '" & r_str_Denomi(r_int_Contad) & "')"
      
      If Not gf_EjecutaSQL(r_str_Cadena, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_HIS_CONLIM.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
      
   Next
   
End Sub

Private Sub fs_GenRpt()

   Dim r_obj_Excel      As Excel.Application
   Dim r_int_ConVer     As Integer
   Dim r_int_ConVar     As Integer
   Dim r_str_TipMon     As String
   
   Screen.MousePointer = 11
   
   'Call fs_GenDat
   Call fs_GeneDB
   
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      .Cells(1, 1) = "SUPERINTENDENCIA DE BANCA, SEGUROS Y AFP"
      .Cells(1, 4) = "REPORTE N�13"
                  
      .Cells(3, 1) = "CONTROL DE L�MITES GLOBALES E INDIVIDUALES APLICABLES A LAS EMPRESAS DEL SISTEMAS FINANCIERO"
      .Cells(4, 1) = "(Contenidas en la Ley N�26702 y normas reglamentarias emitidas por la SBS)"
      .Cells(5, 1) = "Al " & Left(modsec_gf_Fin_Del_Mes("01/" & r_int_PerMes & "/" & ipp_PerAno.Text), 2) & " de " & Left(cmb_PerMes.Text, 1) & LCase(Mid(cmb_PerMes.Text, 2, Len(cmb_PerMes.Text))) & " del " & ipp_PerAno.Text
      .Cells(7, 1) = "EMPRESA:Edpyme Micasita"
      
      .Cells(9, 1) = "I. L�mites Globales (1)"
      .Cells(10, 1) = "Aspecto"
      .Cells(10, 2) = "Base Legal"
      .Cells(10, 3) = "L�mites Legales"
      .Cells(10, 4) = "C�mputo"
      '.Cells(10, 5) = "Numerador"
      '.Cells(10, 6) = "Denominador"
      '.Cells(10, 7) = "Factor"
      '.Cells(10, 8) = "Contraparte"
      
      .Cells(11, 1) = "1. Ratio de Capital Global"
      .Cells(11, 2) = "Art�culo 199� y Vig�simo Cuarta Disposici�n Transitoria de la Ley Gral. Res.SBSN�2115-2009 Res. SBS N�4595-2009, N�4727-2009, N�4729-2009 y N�6328-2009"
      .Cells(11, 3) = "El patrimonio efectivo debe de ser mayor o igualal 10% de la suma de: Activos y Contingentes Ponderados por Riesgo de Cr�dito + 10 multiplicado por Requerimiento " & _
                     "de Patrimonio Efectivo por Riesgo de Mercado + 10 multiplicado por Requerimiento de Patrimonio Efectivo por Riesgo Operacional. (2)"
      .Cells(11, 4) = Format(r_dbl_Evalua(2), "##0.00") & "% de los Activos y Contingentes Ponderadospor Riesgos Totales.%"
      
      .Cells(12, 1) = "2.1. Componentes de Patrimonio B�sico"
      .Cells(12, 2) = "Art�culo 184�, �ltimo p�rrafo literalA Ley Gral. Res. SBS N�4595-2009,  N�4727-2009 y N�4729-2009"
      .Cells(12, 3) = "Los elementos del numeral 3 del literal A del Art.184� de la Ley Gral. s�lo deber�n computarse hasta el 17.65% del monto correspondiente a los componentes considerados " & _
                     "en los numerales 1, 2, 4 y 5 del mismo literal."
      .Cells(12, 4) = Format(r_dbl_Evalua(5), "##0.00") & "% del monto correspondiente a los componentes considerados en los numerales 1, 2, 4 y 5 del literal A del art�culo 184�."
      
      .Cells(13, 1) = "2.2. Patrimonio Suplementario"
      .Cells(13, 2) = "Art�culos 184� y N�185� Ley General Res. SBS N�4595-2009, N�4727-2209 y N�4529-2009"
      .Cells(13, 3) = "El patrimonio suplementario debe de ser menor o igual al 100% del Patrimonio B�sico."
      .Cells(13, 4) = Format(r_dbl_Evalua(8), "##0.00") & "% del patrimonio b�sico de la empresa"
      
      .Cells(14, 1) = "2.3..... Deuda Subordinada Redimible del Patrimonio de Nivel 2"
      .Cells(14, 2) = "Art�culos 184� y N�185� Ley General Res. SBS N�4595-2009, N�4727-2209 y N�4529-2009"
      .Cells(14, 3) = "La deuda subordinada redimible del patrimonio de nivel 2 debe de ser menor o igual al 50% del monto correspondiente a los componentes del patrimonio b�sico considerados " & _
                     "en los numerales 1, 2, 3,4  y 5 del literal A del art�culo 184�."
      .Cells(14, 4) = Format(r_dbl_Evalua(11), "##0.00") & "% de los componentes del patrimonio b�sico de la empresa, considerados en los numerales 1, 2, 3, 4 y 5 del literal A del art�culo 184�."
      
      .Cells(15, 1) = "2.4. Patrimonio Nivel 3"
      .Cells(15, 2) = "Art�culos 184� y N�185� Ley General Res. SBS N�4595-2009, N�4727-2209 y N�4529-2009"
      .Cells(15, 3) = "El patrimonio de nivel 3 debe de ser menor o igual al 250% del monto correspondiente a los componentes del patrimonio b�sico considerados en los numerales 1, 2, 3, 4 y 5 " & _
                     " del literal A del art�culo 184� asignados a cubrir riesgos de mercado."
      .Cells(15, 4) = Format(r_dbl_Evalua(14), "##0.00") & "% de los componentes del patrim�nio b�sico de la empresa, considerados en los numerales 1, 2, 3, 4 y 5 del literal A del art�culo 184� asignado a cubrir riesgos de mercado."
      
      .Cells(16, 1) = "3.1. Riesgos de Mercado - Operaciones con productos financieros derivados"
      .Cells(16, 2) = "Art�culo 200�, num. 3 Ley Gral. Resoluci�n SBS N�1737-2006"
      .Cells(16, 3) = "Hasta el 10% del patrimonio efectivo."
      .Cells(16, 4) = Format(r_dbl_Evalua(17), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(17, 1) = "3.2. Riesgo de Mercado - Tenencias en acciones cotizadas en mecanismos centralizados de negociaci�n, as� como los certificados de participaci�n en fondos mutuos y los certificados " & _
                     " de participaci�n en fondos de inversi�n."
      .Cells(17, 2) = "Art�culo 200�, num. 4 Ley Gral. Resoluci�n SBS N�1067-2005"
      .Cells(17, 3) = "Hasta el 40% del patrimonio efectivo."
      .Cells(17, 4) = Format(r_dbl_Evalua(20), "##0.00") & "% del patrimonioefectivo de la empresa."
      
      .Cells(18, 1) = "4. Requerimiento m�nimo de liquidez en moneda nacional"
      .Cells(18, 2) = "Resoluci�n SBS N�472-2201"
      .Cells(18, 3) = "Activos L�quidos: Mayor o igual al 8% de los pasivos de corto plazo en M.N.(Promedio mensual calendario de saldos diarios)"
      .Cells(18, 4) = Format(r_dbl_Evalua(23), "##0.00") & "% de los pasivos de corto plazo en M.N."
      
      .Cells(19, 1) = "5. Requerimiento m�nimo de liquidez"
      .Cells(19, 2) = "Resoluci�n SBS N�472-2201"
      .Cells(19, 3) = "Activos L�quidos: Mayor o igual al 20% de los pasivosde corto plazo en M.E. (Promedio mensual calendrio de saldos diarios)"
      .Cells(19, 4) = Format(r_dbl_Evalua(26), "##0.00") & "% de los pasivos de corto plazo en M.E."
      
      .Cells(20, 1) = "6. Inversiones en capital social de subsidiarias"
      .Cells(20, 2) = "Art�culo 36�, num. 1 Ley Gral."
      .Cells(20, 3) = "Hasta el 40% del patrimonio contable de la empresa."
      .Cells(20, 4) = Format(r_dbl_Evalua(29), "##0.00") & "% del patrimonio contable de la empresa"

      .Cells(21, 1) = "7. Adquisi�n de facturas mediante factoring"
      .Cells(21, 2) = "Art�culo 200�, num. 1 Ley Gral. Resoluci�n SBS N�1021-98"
      .Cells(21, 3) = "Hasta el 15% del patrimonio efectivo."
      .Cells(21, 4) = Format(r_dbl_Evalua(32), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(22, 1) = "8. Tenencia de oro"
      .Cells(22, 2) = "Art�culo 200�, num. 2 Ley Gral."
      .Cells(22, 3) = "Hasta el 15% del patrimonio efectivo."
      .Cells(22, 4) = Format(r_dbl_Evalua(35), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(23, 1) = "9. Inversi�n en letras hipotecarias de propia emisi�n"
      .Cells(23, 2) = "Circular SBS N�B-1959-94 y similares"
      .Cells(23, 3) = "Hasta el 5% del patrimonio efectivo."
      .Cells(23, 4) = Format(r_dbl_Evalua(38), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(24, 2) = "Circular SBS N�B-1959-94 y similares"
      .Cells(24, 3) = "Excepcionalmente hasta el 10% del patrimonio efectivo previa autorizaci�n de esta Superintendencia, sin exceder el l�mite se�alado en el numeral 7.1 de " & _
                     "Circular SBS N�B-1959-94."
      .Cells(24, 4) = Format(r_dbl_Evalua(41), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(25, 1) = "10. Inversi�n en muebles e inmuebles"
      .Cells(25, 2) = "Art�culo 200�, num. 6 Ley Gral. Resoluci�n SBS N�831-98"
      .Cells(25, 3) = "Hasta el 75% del patrimonio efectivo."
      .Cells(25, 4) = Format(r_dbl_Evalua(44), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(26, 1) = "10.1. Inversi�n en inmuebles"
      .Cells(26, 2) = "Resoluci�n SBS N�831-98"
      .Cells(26, 3) = "Subl�mite 40% del patrimonio efectivo para inversi�n en muebles. (3)"
      .Cells(26, 4) = Format(r_dbl_Evalua(47), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(27, 1) = "10.2. Inversi�n en muebles"
      .Cells(27, 2) = "Resoluci�n SBS N�831-98"
      .Cells(27, 3) = "Subl�mite 40% del patrimonio efectivo para inversi�n en inmuebles. (3)"
      .Cells(27, 4) = Format(r_dbl_Evalua(50), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(28, 1) = "11.1. L�mite la posici�n global de sobreventa de moneda extranjera"
      .Cells(28, 2) = "Resoluci�n SBS N�1455-2003"
      .Cells(28, 3) = "Hasta el 10% del patrimonio efectivo."
      .Cells(28, 4) = Format(r_dbl_Evalua(53), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(29, 1) = "11.2. L�mite la posici�n global de sobrecompra de moneda extranjera"
      .Cells(29, 2) = "Resoluci�n SBS N�1455-2003"
      .Cells(29, 3) = "Hasta el 100% del patrimonio efectivo."
      .Cells(29, 4) = Format(r_dbl_Evalua(56), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(30, 1) = "12. Cr�ditos a directos y trabajadores de la empresa"
      .Cells(30, 2) = "Total cr�ditos a directores y trabajadores de la empresa Art�culo 201� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(30, 3) = "Hasta el 7% del patrimonio efectivo."
      .Cells(30, 4) = Format(r_dbl_Evalua(59), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(31, 1) = "12. Cr�ditos a directos y trabajadores de la empresa"
      .Cells(31, 2) = "Total cr�ditos a directores y trabajadores de la empresa Art�culo 201� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(31, 3) = "Hasta el 7% del patrimonio efectivo."
      .Cells(31, 4) = Format(r_dbl_Evalua(62), "##0.00") & "% del patrimonio efectivo de la empresa."
   
      .Cells(32, 1) = "13. Financiamiento a personas vinculadas a la empresa"
      .Cells(32, 2) = "Art�culo 202� Ley Gral. Res. SBS 445-2000 y N�472-2006 Circular N�B-2148-2005"
      .Cells(32, 3) = "Hasta el 30% del patrimonio efectivo."
      .Cells(32, 4) = Format(r_dbl_Evalua(65), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(33, 1) = "14. Total de financiamientos a soberanos"
      .Cells(33, 2) = "Art�culo 203� de la Ley Gral. Art�culo 212� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(33, 3) = ""
      .Cells(33, 4) = Format(r_dbl_Evalua(68), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(34, 1) = "15. Total de financiamientos a entidades que realizan actividad empresarial del Estado, sin considerar aquellas empresas cuya autonom�a econ�mica y administrativa ha sido declarada por ley."
      .Cells(34, 2) = "Art�culo 203� de la Ley Gral. Art�culo 206� de la Ley Gral. Art�culo 207� de la Ley Gral. Art�culo 208� de la Ley Gral. Art�culo 209� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(34, 3) = "Hasta el 10% del patrimonio efectivo Hasta el 15% del patrimonio efectivo Hasta el 20% del patrimonio efectivo Hasta el 30% del patrimonio efectivo (Sujeto al tipo de garant�a)"
      .Cells(34, 4) = Format(r_dbl_Evalua(71), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      .Cells(35, 1) = "16. Tota de financiamientos otorgados a otras entidades, organismos y dependencias que directa o indirectamente sean considerados o formen parte del Estado Peruano. No se incluyen los financiamientos " & _
                     "se�alados en los numerales 14 y 15 anteriores, ni los otorgados a los gobiernos locales o regionales, ni a COFIDE, AGROBANCO, Fondo MIVIVIENDA, Banco de la Naci�n y Cajas Municipales."
      .Cells(35, 2) = "Art�culo 203� de la Ley Gral. Art�culo 206� de la Ley Gral. Art�culo 207� de la Ley Gral. Art�culo 208� de la Ley Gral. Art�culo 209� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(35, 3) = "Hasta el 10% del patrimonio efectivo Hasta el 15% del patrimonio efectivo Hasta el 20% del patrimonio efectivo Hasta el 30% del patrimonio efectivo (Sujeto al tipo de garant�a)"
      .Cells(35, 4) = Format(r_dbl_Evalua(74), "##0.00") & "% del patrimonio efectivo de la empresa."
      
      For r_int_Contad = 0 To 24 Step 1
         .Range(.Cells(r_int_Contad + 11, 1), .Cells(r_int_Contad + 11, 8)).WrapText = True
      Next
      
      .Range(.Cells(1, 8), .Cells(1, 4)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(1, 1), .Cells(3, 8)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(5, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 1)).Font.Bold = True
                  
      .Range(.Cells(1, 1), .Cells(1, 2)).Merge
      .Range(.Cells(3, 1), .Cells(3, 4)).Merge
      .Range(.Cells(4, 1), .Cells(4, 4)).Merge
      .Range(.Cells(5, 1), .Cells(5, 4)).Merge
      .Range(.Cells(23, 1), .Cells(24, 1)).Merge
      .Range(.Cells(9, 1), .Cells(9, 4)).Merge
                  
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(10, 1), .Cells(10, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(9, 1), .Cells(35, 8)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(9, 1), .Cells(35, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                 
      .Range(.Cells(9, 1), .Cells(10, 8)).Font.Bold = True
      
      
      .Cells(37, 1) = "II. L�mites Individuales(1)"
      .Cells(38, 1) = "Aspecto (4)"
      .Cells(38, 2) = "Base Legal"
      .Cells(38, 3) = "L�mites Legales"
      .Cells(38, 4) = "C�mputo (6)"
      '.Cells(38, 5) = "Numerador"
      '.Cells(38, 6) = "Denominador"
      '.Cells(38, 7) = "Factor"
      '.Cells(38, 8) = "Contraparte"
      
      .Cells(39, 1) = "1. Total de financiamiento a un deudor o grupo de deudores que representa riesgo �nico (5)."
      .Cells(39, 2) = "Art. 203�; 204�; 205�; 206�; 207�; 208�; 209�; 210�; 211� y 212� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(39, 3) = "Hasta el 30% del patrimonio efectivo debiendose adem�s tener encuenta los subl�mites contemplados en el numeral 4 de las Circular N�B-2148-2005. Hasta el 50%, si el exceso se encuentra representado " & _
                     "por cartas de cr�dito de empresas del sistema financiero del exterior de conformidad con elnumeral 4 del Art. 205� de la Ley Gral."
      .Cells(39, 4) = "1)" & Format(r_dbl_Evalua(77), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo de deudores que representa riesgo �nico, con el mayor financiamiento."
      .Cells(40, 4) = "2)" & Format(r_dbl_Evalua(80), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo de deudores que representa riesgo �nico, con el segundo mayor financiamiento."
      .Cells(41, 4) = "3)" & Format(r_dbl_Evalua(83), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo de deudores que representa riesgo �nico, con el tercer mayor financiamiento."
      
      .Cells(42, 1) = "2. Financiamiento directo e indirecto a empresas del sistema  financiero establecidas en el pa�s, dep�sitos constituidos en ellas, avales, fianzas y otras garant�as que se haya recibido de dichas empresas."
      .Cells(42, 2) = "Art�culo 204�; p�rrafo 1 Ley Gral. Art�culo 203� Ley Gral. Art. 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(42, 3) = "Hasta el 30% del patrimonio efectivo."
      .Cells(42, 4) = "1)" & Format(r_dbl_Evalua(86), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(43, 4) = "2)" & Format(r_dbl_Evalua(89), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(44, 4) = "3)" & Format(r_dbl_Evalua(92), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(45, 1) = "3. Financiamiento directo e indirecto a empresas bancarias o financieras del exterior, los dep�sitos constituidos en ellas y los avales avales, fianzas y otras garant�as que se haya recibido de dichas instituciones."
      .Cells(45, 2) = "Art�culo 204�; p�rrafo 1 Ley Gral. Art�culo 203� Ley Gral. Art. 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(45, 3) = "Hasta el 5% del patrimonio efectivo, en caso de empresas no sujetas a supervisi�n por organismos similares a la sbs. Hasta el 50% del patrimonio efectivo, siempre que el exceso se encuentre representado por cartas de cr�dito, incluyendo la modalidad de stand by letter of credit."
      .Cells(45, 4) = "1)" & Format(r_dbl_Evalua(95), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(46, 4) = "2)" & Format(r_dbl_Evalua(98), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(47, 4) = "3)" & Format(r_dbl_Evalua(101), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(48, 2) = "Art�culo 205�; numeral 2 y 4, 203� y 212� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(48, 3) = "Se puede exceder el l�mite anterior de 5% hasta el 10% del patrimonio efectivo, en caso de empresas sujetas a supervisi�n por organismos similares a ls SBS, y no son bancos de 1ra categor�a. Hasta el 50% del patrimonio efectivo, siempre que el exceso se encuentre representado por cartas de cr�dito, incluyendo la modalidad de stand by letter of credit."
      .Cells(48, 4) = "1)" & Format(r_dbl_Evalua(104), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(49, 4) = "2)" & Format(r_dbl_Evalua(107), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(50, 4) = "3)" & Format(r_dbl_Evalua(110), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(51, 2) = "Art�culo 205�; numeral 3 y 4, 203� y 212� de la Ley Gral. Circular N�B-2148-2005"
      .Cells(51, 3) = "Se puede exceder el l�mite anterior de 5% y 10% anteriores hasta el 30% del patrimonio efectivo en caso de bancos de 1ra categor�a. Hasta el 50% del patrimonio efectivo, siempre que el exceso se encuentre representado por cartas de cr�dito, incluyendo la modalidad de stand by letter of credit."
      .Cells(51, 4) = "1)" & Format(r_dbl_Evalua(113), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(52, 4) = "2)" & Format(r_dbl_Evalua(116), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(53, 4) = "3)" & Format(r_dbl_Evalua(119), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(54, 1) = "4. Financiamiento mediante cr�ditos, inversiones y contingentes a favor de una misma persona natural jur�dica directa o indirectamente (para residentes en el pa�s con excepci�n de empresas del sistema financiero y de las entidades consideradas en los numerales 14, 15 y 16 de la selecci�n L�mites Globales)"
      .Cells(54, 2) = "Art�culo 206� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(54, 3) = "Hasta el 10% del patrimonio efectivo."
      .Cells(54, 4) = "1)" & Format(r_dbl_Evalua(122), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(55, 4) = "2)" & Format(r_dbl_Evalua(125), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(56, 4) = "3)" & Format(r_dbl_Evalua(128), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(57, 2) = "Art�culo 207� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(57, 3) = "Se puede exceder el l�mite contemplado en el primer p�rrafo del art�culo 206� de la Ley General hasta el 15% del patrimonio efectivo (sujeto al tipo de garant�a)"
      .Cells(57, 4) = "1)" & Format(r_dbl_Evalua(131), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(58, 4) = "2)" & Format(r_dbl_Evalua(134), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(59, 4) = "3)" & Format(r_dbl_Evalua(137), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(60, 2) = "Art�culo 208� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(60, 3) = "Se puede exceder los l�mite contemplados en el primer p�rrafo del Art. 206� y en el Art. 207� de la Ley General hasta el 20% del patrimonio efectivo (sujeto al tipo de garant�a)"
      .Cells(60, 4) = "1)" & Format(r_dbl_Evalua(140), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(61, 4) = "2)" & Format(r_dbl_Evalua(143), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(62, 4) = "3)" & Format(r_dbl_Evalua(146), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(63, 2) = "Art�culo 209� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(63, 3) = "Se puede exceder los l�mite contemplados en el primer p�rrafo del Art. 206�; en el Art. 207� y en el Art. 208� de la Ley General hasta el 30% del patrimonio efectivo (sujeto al tipo de garant�a)"
      .Cells(63, 4) = "1)" & Format(r_dbl_Evalua(149), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(64, 4) = "2)" & Format(r_dbl_Evalua(152), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(65, 4) = "3)" & Format(r_dbl_Evalua(155), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."

      .Cells(66, 1) = "5. Cr�ditos, contingentes, inversiones y arrendatarios financieros otorgados a una misma persona natural o jur�dica residentes en el exterior (conexcepcion de empresas del sistema financiero)"
      .Cells(66, 2) = "Art�culo 211� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(66, 3) = "Hasta el 5% del patrimonio efectivo."
      .Cells(66, 4) = "1)" & Format(r_dbl_Evalua(158), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(67, 4) = "2)" & Format(r_dbl_Evalua(161), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(68, 4) = "3)" & Format(r_dbl_Evalua(164), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."

      .Cells(69, 2) = "Art�culo 211� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(69, 3) = "Se puede exceder los l�mites anteriores de 5% y 10% hasta el 30% del patrimonio efectivo (Sujeto al tipo de garant�a)."
      .Cells(69, 4) = "1)" & Format(r_dbl_Evalua(167), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(70, 4) = "2)" & Format(r_dbl_Evalua(170), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(71, 4) = "3)" & Format(r_dbl_Evalua(173), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(72, 1) = "6. Fianzas otorgadas para garantizar la suscripci�n de contratos derivados de procesos de licitaci�n p�blica."
      .Cells(72, 2) = "Art�culo 206� Ley Gral. Art�culo 203� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(72, 3) = "Hasta el 30% del patrimonio efectivo."
      .Cells(72, 4) = "1)" & Format(r_dbl_Evalua(176), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(73, 4) = "2)" & Format(r_dbl_Evalua(179), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(74, 4) = "3)" & Format(r_dbl_Evalua(182), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(75, 1) = "7. Inversiones en letras hipotecarias de un mismo emisor."
      .Cells(75, 2) = "Circular N�B-1959-94 y similares"
      .Cells(75, 3) = "Hasta el 10% del patrimonio efectivo."
      .Cells(75, 4) = "1)" & Format(r_dbl_Evalua(185), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya adquirido del emisor con el mayor financiamiento."
      .Cells(76, 4) = "2)" & Format(r_dbl_Evalua(188), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya adquirido del emisor con el segundo mayor financiamiento."
      .Cells(77, 4) = "3)" & Format(r_dbl_Evalua(191), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya adquirido del emisor con el tercer mayor financiamiento."
      
      .Cells(78, 1) = "8. Cr�dito a directores y trabajadores de la empresa."
      .Cells(78, 2) = "Art�culo 201� Ley Gral. Art�culo 212� Ley Gral. Circular N�B-2148-2005"
      .Cells(78, 3) = "Hasta el 5% de lo se�alado en el numeral 12 de la secci�n L�mites Globales (es decir 0.35% del patrimonio efectivo)."
      .Cells(78, 4) = "1)" & Format(r_dbl_Evalua(194), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el mayor financiamiento."
      .Cells(79, 4) = "2)" & Format(r_dbl_Evalua(197), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el segundo mayor financiamiento."
      .Cells(80, 4) = "3)" & Format(r_dbl_Evalua(200), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya otorgado a la persona o grupo con el tercer mayor financiamiento."
      
      .Cells(81, 1) = "9. Warrants recibidos en garant�a."
      .Cells(81, 2) = "Art�culo 204�; p�rrafo 2 Ley Gral. Circular N�B-2148-2005"
      .Cells(81, 3) = "Hasta el 60% del patrimonio efectivo."
      .Cells(81, 4) = "1)" & Format(r_dbl_Evalua(203), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del AGD con el mayor importe."
      .Cells(82, 4) = "2)" & Format(r_dbl_Evalua(206), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del AGD con el segundo mayor importe."
      .Cells(83, 4) = "3)" & Format(r_dbl_Evalua(209), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del AGD con el tercer mayor importe."
      
      .Cells(84, 1) = "10. Coberturas que otorguen un patrimonio aut�nomo de seguro de cr�dito o un fondo de garant�a creado por el Estado a favor de una misma empresa."
      .Cells(84, 2) = "Art�culo 204�; p�rrafo 3 Ley Gral. Art�culo 212� Ley Gral."
      .Cells(84, 3) = ""
      .Cells(84, 4) = "1)" & Format(r_dbl_Evalua(212), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del patrimonio aut�nomo o fondo con la mayor cobertura otorgada."
      .Cells(85, 4) = "2)" & Format(r_dbl_Evalua(215), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del patrimonio aut�nomo o fondo con la segundo mayor cobertura otorgada."
      .Cells(86, 4) = "3)" & Format(r_dbl_Evalua(218), "##0.00") & "% del patrimonio efectivo de la empresa, que se haya recibido del patrimonio aut�nomo o fondo con la tercera mayor cobertura otorgada."
      
      .Cells(87, 1) = "11. Inversi�n en el capital social de una subsidiaria."
      .Cells(87, 2) = "Art�culo 36�; num. 2 Ley Gral."
      .Cells(87, 3) = "Mayor o igual a 3/5 partes del capital social de la subsidiaria."
      .Cells(87, 4) = Format(r_dbl_Evalua(221), "##0.00") & " partes del capital social de la subsidiaria."
      .Cells(88, 4) = Format(r_dbl_Evalua(224), "##0.00") & "Calcular el l�mite para cada subsidiaria."
      
      For r_int_Contad = 39 To 88 Step 1
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 8)).WrapText = True
      Next
      
      For r_int_Contad = 39 To 87 Step 3
         If r_int_Contad = 45 Then
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 8, 1)).Merge
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 8, 1)).WrapText = True
            r_int_Contad = r_int_Contad + 6
         ElseIf r_int_Contad = 54 Then
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 11, 1)).Merge
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 11, 1)).WrapText = True
            r_int_Contad = r_int_Contad + 9
         ElseIf r_int_Contad = 66 Then
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 5, 1)).Merge
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 5, 1)).WrapText = True
            r_int_Contad = r_int_Contad + 3
         ElseIf r_int_Contad <> 87 Then
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 2, 1)).Merge
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 2, 1)).WrapText = True
         Else
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 1, 1)).Merge
            .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad + 1, 1)).WrapText = True
         End If
      Next
            
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(38, 1), .Cells(38, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(37, 1), .Cells(95, 8)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(37, 1), .Cells(88, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                 
      .Range(.Cells(37, 1), .Cells(38, 8)).Font.Bold = True
      
      For r_int_Contad = 40 To 88 Step 3
         .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 8)).Borders(xlEdgeTop).LineStyle = xlNone
         If r_int_Contad <> 88 Then
            .Range(.Cells(r_int_Contad, 2), .Cells(r_int_Contad, 8)).Borders(xlEdgeBottom).LineStyle = xlNone
         End If
      Next
      
      .Range(.Cells(90, 1), .Cells(90, 4)).Merge
      .Range(.Cells(91, 1), .Cells(91, 4)).Merge
      .Range(.Cells(92, 1), .Cells(92, 4)).Merge
      .Range(.Cells(93, 1), .Cells(93, 4)).Merge
      .Range(.Cells(94, 1), .Cells(94, 4)).Merge
      .Range(.Cells(95, 1), .Cells(95, 4)).Merge
      
      .Range(.Cells(90, 1), .Cells(90, 1)).RowHeight = 30
      '.Range(.Cells(92, 1), .Cells(92, 1)).RowHeight = 30
      .Range(.Cells(93, 1), .Cells(93, 1)).RowHeight = 30
      .Range(.Cells(94, 1), .Cells(94, 1)).RowHeight = 45
      
      .Range(.Cells(90, 1), .Cells(90, 1)).WrapText = True
      .Range(.Cells(92, 1), .Cells(92, 1)).WrapText = True
      .Range(.Cells(93, 1), .Cells(93, 1)).WrapText = True
      .Range(.Cells(94, 1), .Cells(94, 1)).WrapText = True
            
      .Cells(90, 1) = "(1) El patrimonio efectivo que deber� emplearse para el c�mputo de los l�mites ser� el �ltimo remitido por la empresa y validado por esta Superintendencia, salvo para el caso " & _
                     "del ratio de capital global y los subl�mites de los componentes del patrimonio efectivo."
      .Cells(91, 1) = "(2) Se deber� considerar el cronograma de adecuaci�n se�alado en la Vig�simo Cuarta Disposici�n Transitoria de la Ley General."
      .Cells(92, 1) = "(3) Se deber� considerar los saldos netos de depresaci�n y p�rdida por deterioro acumuladas. No se incluyen los bienes dados en arrendamiento financiero, bienes recuperados, " & _
                     "bienes adjudicados y recibidos en pago."
      .Cells(93, 1) = "(4) Para cada l�mite y/o subl�mite individual indicar tres (3) mayores exposiciones. Por ejemplo: para la medici�n del l�mite individual a que se refiere el art�culo 206�, " & _
                     "la empresa deber� reportar los tres (3) mayores financiamientos otorgados sin garant�as al deudor o grupo de deudores domiciliados en el pa�s qu representan riesgo �nico."
      .Cells(94, 1) = "(5) Las empresasdeber�naplicarlos l�mites de concentraci�n a que se refiere en los art�culos 204� al 211� de la Ley General considerando el criterio de riesgo �nico de conformidad " & _
                     "de conformidad con los dispuesto en el art�culo 203� de la Ley General y en el cap�tulo II de las Normas Especiales sobre Vinculaci�n y Grupo Econ�mico, de tal forma que un grupo " & _
                     "de contrapartes relacionadas que representen riesgo �nico no podr� exceder del treinta por ciento (30%) del patrimonio efectivo de la empresa, de conformidad con lo establecido en " & _
                     "el numeral 4 de la Circular B-2148-2005 y modificatorias."
      .Cells(95, 1) = "(6) Luego de reportar los indicadores de exposici�n se deber� incluir el nombre de la contraparte (persona o grupo, AGD,subsidiaria, patrimonio aut�nomo o fondo de garant�a)."
      
      
      .Cells(100, 2) = "_________________________"
      .Cells(101, 2) = "Gerente General"
      
      .Cells(100, 3) = "______________________________"
      .Cells(101, 3) = "Gerente de Unidad de Riesgos"
      
      .Range(.Cells(100, 2), .Cells(101, 3)).HorizontalAlignment = xlHAlignCenter
                 
                                              
      .Range(.Cells(1, 1), .Cells(110, 110)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(110, 110)).Font.Size = 9
            
      .Columns("A").ColumnWidth = 49
      .Columns("B").ColumnWidth = 49
      .Columns("C").ColumnWidth = 49
      .Columns("D").ColumnWidth = 49
      .Columns("E").ColumnWidth = 15
      .Columns("F").ColumnWidth = 15
      .Columns("G").ColumnWidth = 15
      .Columns("H").ColumnWidth = 15
                  
      .Columns("E:H").NumberFormat = "###,###,##0.00"
                         
   End With
   
   'Bloquear el archivo
   r_obj_Excel.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="382-6655"
   
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing

             
End Sub


Private Sub fs_GenDat_DB()

   r_int_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   Erase r_str_Denomi
   Erase r_dbl_Evalua()
   
   g_str_Parame = "SELECT * FROM HIS_CONLIM WHERE "
   g_str_Parame = g_str_Parame & "CONLIM_PERMES = " & r_int_PerMes & " AND "
   g_str_Parame = g_str_Parame & "CONLIM_PERANO = " & r_int_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY CONLIM_NUMITE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
               
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
   
      g_rst_Princi.MoveFirst
      r_int_ConAux = -1
      r_int_ConTem = -1
           
      Do While Not g_rst_Princi.EOF
                     
         r_dbl_Evalua(r_int_ConAux + 1) = g_rst_Princi!CONLIM_MUMERA
         r_dbl_Evalua(r_int_ConAux + 2) = g_rst_Princi!CONLIM_DENOMI
         r_dbl_Evalua(r_int_ConAux + 3) = g_rst_Princi!CONLIM_COMPUT
         r_str_Denomi(r_int_ConTem + 1) = IIf(IsNull(g_rst_Princi!CONLIM_CONTRA) = True, "", Trim(g_rst_Princi!CONLIM_CONTRA))
         
         r_int_ConAux = r_int_ConAux + 3
         r_int_ConTem = r_int_ConTem + 1
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   
   End If
  
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
     
         
End Sub








