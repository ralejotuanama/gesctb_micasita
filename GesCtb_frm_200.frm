VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_InvDpf_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   Icon            =   "GesCtb_frm_200.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   2625
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   5
         Top             =   60
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
            Height          =   480
            Left            =   600
            TabIndex        =   6
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Generar Asientos - Interés Devengados"
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
            Picture         =   "GesCtb_frm_200.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1095
         Left            =   30
         TabIndex        =   7
         Top             =   1470
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
         Begin VB.ComboBox cmb_CodMes 
            Height          =   315
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1650
            TabIndex        =   1
            Top             =   600
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
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   300
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   300
            TabIndex        =   9
            Top             =   615
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   780
         Width           =   7365
         _Version        =   65536
         _ExtentX        =   12991
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
         Begin VB.CommandButton cmd_Generar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_200.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Generar Asientos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6750
            Picture         =   "GesCtb_frm_200.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_InvDpf_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_PerMes    As Integer
Dim l_int_PerAno    As Integer
Dim l_str_FecDia    As String
Dim l_str_FecPer    As String
Dim l_int_Conteo    As Integer
Dim l_str_Codigo    As String

Private Sub cmb_CodMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub cmd_Generar_Click()
Dim r_dbl_TipSbs As Double
Dim r_bol_Estado As Boolean
Dim r_str_CtaDeb As String
Dim r_str_CtaHab As String
Dim r_str_CadAux As String
Dim r_int_Contad As Integer

   l_int_PerMes = CInt(cmb_CodMes.ItemData(cmb_CodMes.ListIndex))
   l_int_PerAno = ipp_PerAno.Text
   l_str_FecDia = Format(ff_Ultimo_Dia_Mes(l_int_PerMes, l_int_PerAno), "00")
   l_str_FecPer = l_str_FecDia & "/" & Format(l_int_PerMes, "00") & "/" & l_int_PerAno
   
   Call fs_ValGeneracion(Format(l_str_FecPer, "yyyymmdd"), l_int_Conteo, l_str_Codigo)
   If l_int_Conteo > 0 Then
      MsgBox "El Período " & Format(l_str_FecPer, "yyyy-mm") & " ya fue generado", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
'   'TIPO CAMBIO SBS(2) - VENTA(1)
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, Format(l_str_FecPer, "yyyymmdd"), 1)
   If r_dbl_TipSbs = 0 Then
      MsgBox "No hay tipo de cambio SBS ingresado para la fecha (" & l_str_FecPer & ")", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Valida si existen la dinamica contable para cada cuenta
   r_bol_Estado = True
   r_str_CadAux = ""
   For r_int_Contad = 0 To frm_Ctb_InvDpf_01.grd_Listad.Rows - 1
       If CLng(Format(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 8), "yyyymmdd")) >= CLng(Format(l_str_FecPer, "yyyymmdd")) Then
          'buscar las cuenta
          r_str_CtaDeb = "": r_str_CtaHab = ""
          Call fs_BuscarCtas(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 15), frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 16), _
                             frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 17), r_str_CtaDeb, r_str_CtaHab)
          If r_str_CtaDeb = "" Or r_str_CtaHab = "" Then
             r_bol_Estado = False
             r_str_CadAux = r_str_CadAux & " - " & frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 0)
          End If
       End If
   Next
   
   If r_bol_Estado = False Then
      MsgBox "Falta definir la dinamica contable del rubro devengados." & vbCrLf & "Nro Cuentas: " & r_str_CadAux, vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Este proceso se genera una vez al mes, Está seguro de generalo?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de generar los asientos contables?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GeneraAsiento
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_CodMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_CodMes, 1, "033")
      
   ipp_PerAno = Mid(date, 7, 4)
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
 
   Call gs_BuscarCombo_Item(cmb_CodMes, r_int_PerMes)
   ipp_PerAno.Text = Format(r_int_PerAno, "0000")
End Sub

Public Sub fs_ValGeneracion(p_FecPer As String, ByRef p_Conteo As Integer, ByRef p_Codigo As String)

   p_Conteo = 0
   p_Codigo = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT (SELECT COUNT(*)  "
   g_str_Parame = g_str_Parame & "          FROM MNT_PARDES A  "
   g_str_Parame = g_str_Parame & "         WHERE A.PARDES_CODGRP = 129  "
   g_str_Parame = g_str_Parame & "           AND PARDES_SITUAC = '1'  "
   g_str_Parame = g_str_Parame & "           AND SUBSTR(TRIM(PARDES_DESCRI),0,8) = '" & p_FecPer & "') AS CONTEO,  "
   g_str_Parame = g_str_Parame & "       NVL((SELECT MAX(TO_NUMBER(PARDES_CODITE))  "
   g_str_Parame = g_str_Parame & "              FROM MNT_PARDES A  "
   g_str_Parame = g_str_Parame & "             WHERE PARDES_CODITE <> '000000'  "
   g_str_Parame = g_str_Parame & "               AND A.PARDES_CODGRP = 129),0)+1 AS CODIGO  "
   g_str_Parame = g_str_Parame & "  FROM DUAL  "
 
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   
   p_Conteo = g_rst_Genera!CONTEO
   p_Codigo = Format(g_rst_Genera!CODIGO, "000000")
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Public Sub fs_GeneraAsiento()
Dim r_arr_LogPro()    As modprc_g_tpo_LogPro
Dim r_str_Origen      As String
Dim r_str_TipNot      As String
Dim r_int_NumLib      As Integer
Dim r_str_AsiGen      As String
Dim r_int_NumAsi      As Integer
Dim r_str_Glosa       As String
Dim r_dbl_Import      As Double
Dim r_dbl_MtoSol      As Double
Dim r_dbl_MtoDol      As Double
Dim r_str_DebHab      As String
Dim r_dbl_TipSbs      As Double
Dim r_str_CtaHab      As String
Dim r_str_CtaDeb      As String
    
Dim r_int_NumAsi_2    As Integer
Dim r_str_FecPer_2    As String
Dim r_int_PerMes_2    As Integer
Dim r_int_PerAno_2    As Integer
Dim r_int_Contad      As Integer
Dim r_str_ImpAux      As String
Dim r_str_AsiErr_A    As String
Dim r_str_AsiErr_B    As String
Dim r_str_CadAux      As String
Dim r_str_FecPrPgoC   As String
Dim r_str_FecPrPgoL   As String
Dim r_str_AsiGen_2    As String
Dim r_str_FecPrPgoC_2 As String
Dim r_str_FecPrPgoL_2 As String

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 12
   r_str_AsiGen = ""
   r_str_CtaHab = ""
   r_str_CtaDeb = ""
      
   'Inicializa variables
   r_int_NumAsi = 0
   r_str_FecPrPgoC = Format(l_str_FecPer, "yyyymmdd")
   r_str_FecPrPgoL = l_str_FecPer
   
   r_str_AsiGen_2 = ""
   r_str_FecPer_2 = DateAdd("d", 1, l_str_FecPer)
   r_int_PerMes_2 = CInt(Format(r_str_FecPer_2, "mm"))
   r_int_PerAno_2 = CInt(Format(r_str_FecPer_2, "yyyy"))
   r_str_FecPrPgoC_2 = Format(r_str_FecPer_2, "yyyymmdd")
   r_str_FecPrPgoL_2 = r_str_FecPer_2
            
'   'TIPO CAMBIO SBS(2) - VENTA(1)
   r_dbl_TipSbs = moddat_gf_ObtieneTipCamDia(2, 2, Format(l_str_FecPer, "yyyymmdd"), 1)
      
   For r_int_Contad = 0 To frm_Ctb_InvDpf_01.grd_Listad.Rows - 1
       'If frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 12) = 1 Then 'solo vigentes
         'Sólo realizará el cálculo del devengado para los depósitos cuya fecha de vencimiento sea mayor o igual al último día del mes seleccionado.
         'If CLng(Format(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 8), "yyyymmdd")) <= CLng(Format(l_str_FecPer, "yyyymmdd")) And _
         '   CLng(Format(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 8), "yyyymmdd")) >= CLng(l_int_PerAno & Format(l_int_PerMes, "00") & "01") Then
         If CLng(Format(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 8), "yyyymmdd")) >= CLng(Format(l_str_FecPer, "yyyymmdd")) Then
         
            r_str_Glosa = Mid(Trim("OPERACION DPF - INTDEV - " & frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 0)), 1, 60)
            'Obteniendo Nro. de Asiento
            r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, l_int_PerAno, l_int_PerMes, r_str_Origen, r_int_NumLib)
            r_str_AsiGen = r_str_AsiGen & " - " & CStr(r_int_NumAsi)
            'Insertar en cabecera
             Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                  r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL, "1")
            'cálculo devengado
            r_dbl_Import = 0: r_str_ImpAux = ""
            Call frm_Ctb_InvDpf_01.fs_CalDevengado(l_str_FecPer, Format(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 8), "yyyymmdd"), _
                                   frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 9), Format(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 7), "yyyymmdd"), _
                                   frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 4), frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 6), r_str_ImpAux)
            r_dbl_Import = CDbl(r_str_ImpAux)
            
            r_dbl_MtoSol = 0: r_dbl_MtoDol = 0
            
            If frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 17) = 1 Then
               'Entrega a rendir Soles:
                r_dbl_MtoSol = r_dbl_Import
                r_dbl_MtoDol = Format(CDbl(r_dbl_Import / r_dbl_TipSbs), "###,###,##0.00")
            Else 'Entrega a rendir Dolares:
                r_dbl_MtoSol = Format(CDbl(r_dbl_Import * r_dbl_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
                r_dbl_MtoDol = r_dbl_Import
            End If
                        
            'buscar las cuenta
            Call fs_BuscarCtas(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 15), frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 16), _
                               frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 17), r_str_CtaDeb, r_str_CtaHab)
            
            'Insertar en detalle
            If r_str_CtaDeb = "" Or r_str_CtaHab = "" Then
               r_str_AsiErr_A = r_str_AsiErr_A & " - " & CStr(r_int_NumAsi)
            End If
            If r_str_CtaDeb <> "" Then
            Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                                 r_int_NumAsi, 1, r_str_CtaDeb, CDate(r_str_FecPrPgoL), r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
            End If
            If r_str_CtaHab <> "" Then
               Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, l_int_PerAno, l_int_PerMes, r_int_NumLib, _
                                                    r_int_NumAsi, 2, r_str_CtaHab, CDate(r_str_FecPrPgoL), r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL))
            End If
            '-------------------------------------------------------------------------------------------------------------------------
            'Avanzar un Dia Al Periodo
            '-------------------------------------------------------------------------------------------------------------------------
            'Obteniendo Nro. de Asiento
            r_int_NumAsi_2 = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno_2, r_int_PerMes_2, r_str_Origen, r_int_NumLib)
            r_str_AsiGen_2 = r_str_AsiGen_2 & " - " & CStr(r_int_NumAsi_2)
            'Insertar en cabecera
            Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno_2, r_int_PerMes_2, r_int_NumLib, _
                 r_int_NumAsi_2, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FecPrPgoL_2, "1")
            'cálculo devengado
            r_dbl_Import = 0: r_str_ImpAux = ""
            Call frm_Ctb_InvDpf_01.fs_CalDevengado(r_str_FecPer_2, Format(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 8), "yyyymmdd"), _
                                   frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 9), Format(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 7), "yyyymmdd"), _
                                   frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 4), frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 6), r_str_ImpAux)
            r_dbl_Import = CDbl(r_str_ImpAux)
            
            r_dbl_MtoSol = 0: r_dbl_MtoDol = 0
            If frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 17) = 1 Then
               'Entrega a rendir Soles:
                r_dbl_MtoSol = r_dbl_Import
                r_dbl_MtoDol = Format(CDbl(r_dbl_Import / r_dbl_TipSbs), "###,###,##0.00")
            Else 'Entrega a rendir Dolares:
                r_dbl_MtoSol = Format(CDbl(r_dbl_Import * r_dbl_TipSbs), "###,###,##0.00") 'Importe * CONVERTIDO
                r_dbl_MtoDol = r_dbl_Import
            End If
            'buscar las cuenta
            Call fs_BuscarCtas(frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 15), frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 16), _
                               frm_Ctb_InvDpf_01.grd_Listad.TextMatrix(r_int_Contad, 17), r_str_CtaDeb, r_str_CtaHab)
            
            'Insertar en detalle
            If r_str_CtaDeb = "" Or r_str_CtaHab = "" Then
               r_str_AsiErr_B = r_str_AsiErr_B & " - " & CStr(r_int_NumAsi_2)
            End If
            If r_str_CtaDeb <> "" Then
               Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno_2, r_int_PerMes_2, r_int_NumLib, _
                                                    r_int_NumAsi_2, 1, r_str_CtaDeb, CDate(r_str_FecPrPgoL_2), r_str_Glosa, "D", _
                                                    r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL_2))
            End If
            If r_str_CtaHab <> "" Then
               Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno_2, r_int_PerMes_2, r_int_NumLib, _
                                                    r_int_NumAsi_2, 2, r_str_CtaHab, CDate(r_str_FecPrPgoL_2), r_str_Glosa, "H", _
                                                    r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecPrPgoL_2))
            End If
         End If
       'End If
   Next
               
      
   r_str_CadAux = Left(Format(l_str_FecPer, "yyyymmdd") & " - Asientos(" & Trim(r_str_AsiGen) & ") (" & Trim(r_str_AsiGen_2) & ")", 120)
   
   g_str_Parame = "USP_INSERTA_MNT_PARDES ("
   g_str_Parame = g_str_Parame & "'129', "
   g_str_Parame = g_str_Parame & "'" & l_str_Codigo & "' , "
   g_str_Parame = g_str_Parame & "'" & r_str_CadAux & "', "
   g_str_Parame = g_str_Parame & "1, "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      Exit Sub
   End If
         
   MsgBox "Se culminó proceso de generación de asientos contables:" & vbCrLf & _
          "Asientos Fecha ( " & Format(l_str_FecPer, "dd/mm/yyyy") & " ) " & vbCrLf & _
          "Asientos Generados: " & Trim(r_str_AsiGen) & vbCrLf & _
          "Asientos Errados: " & r_str_AsiErr_A & vbCrLf & " " & vbCrLf & _
          "Asientos Fecha ( " & Format(r_str_FecPer_2, "dd/mm/yyyy") & " ) " & vbCrLf & _
          "Asientos Generados: " & Trim(r_str_AsiGen_2) & vbCrLf & _
          "Asientos Errados: " & r_str_AsiErr_B, vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Function fs_BuscarCtas(p_CODENT_DES As Integer, p_CODENT_ORI As Integer, p_CODMON As Integer, ByRef p_CtaDeb As String, ByRef p_CtaHab As String)
   'extrae el numero de cuenta
   p_CtaDeb = ""
   p_CtaHab = ""
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT A.CTADPF_CTADEB_01, A.CTADPF_CTADEB_02, A.CTADPF_CTAHAB_01, A.CTADPF_CTAHAB_02  "
   g_str_Parame = g_str_Parame & "    FROM CTB_CTADPF A  "
   g_str_Parame = g_str_Parame & "   WHERE A.CTADPF_CODENT_DES =  " & p_CODENT_DES
   g_str_Parame = g_str_Parame & "     AND A.CTADPF_TIPDPF = 6  "
   g_str_Parame = g_str_Parame & "     AND A.CTADPF_CODENT_ORI = " & p_CODENT_ORI
   g_str_Parame = g_str_Parame & "     AND A.CTADPF_CODMON = " & p_CODMON
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Function
   End If
            
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ninguna cuenta contable para generar el asiento", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Function
   Else
      p_CtaDeb = Trim(g_rst_Princi!CTADPF_CTADEB_01 & "")
      p_CtaHab = Trim(g_rst_Princi!CTADPF_CTAHAB_01 & "")
   End If
End Function

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Generar)
   End If
End Sub
