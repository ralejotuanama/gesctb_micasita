VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2850
   ClientLeft      =   15330
   ClientTop       =   3375
   ClientWidth     =   4200
   Icon            =   "GesCtb_frm_806.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2895
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   5106
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
         Width           =   4125
         _Version        =   65536
         _ExtentX        =   7276
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
            TabIndex        =   7
            Top             =   120
            Width           =   3375
            _Version        =   65536
            _ExtentX        =   5953
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Reporte Crediticio de Deudores (RCD)"
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
            Picture         =   "GesCtb_frm_806.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   780
         Width           =   4125
         _Version        =   65536
         _ExtentX        =   7276
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
            Left            =   630
            Picture         =   "GesCtb_frm_806.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Archivo Texto"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   3510
            Picture         =   "GesCtb_frm_806.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_806.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Procesar RCD"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1650
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
         Height          =   885
         Left            =   30
         TabIndex        =   9
         Top             =   1470
         Width           =   4125
         _Version        =   65536
         _ExtentX        =   7276
         _ExtentY        =   1561
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
         Begin VB.ComboBox cmb_Period 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   990
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
            Left            =   150
            TabIndex        =   11
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   150
            TabIndex        =   10
            Top             =   120
            Width           =   465
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   12
         Top             =   2370
         Width           =   4125
         _Version        =   65536
         _ExtentX        =   7276
         _ExtentY        =   767
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
         Begin Threed.SSPanel pnl_BarPro 
            Height          =   315
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7064
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SSPanel2"
            ForeColor       =   16777215
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
            FloodType       =   1
            FloodColor      =   49152
            Font3D          =   2
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_FecIni        As String
Dim l_str_FecFin        As String
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Dim r_lng_NumReg        As Long
Dim r_lng_TotReg        As Long

Private Sub cmd_ExpArc_Click()

   If cmb_Period.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Period)
      Exit Sub
   End If
   
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de generar el archivo?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   l_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00") & "01"
   l_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_Period.ItemData(cmb_Period.ListIndex), CInt(ipp_PerAno.Text)), "00")
        
   Screen.MousePointer = 11
      
      Call fs_Genera_ArcRcd("RCD", l_str_FecIni, l_str_FecFin)
            
   Screen.MousePointer = 0
   
   MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   
End Sub

Private Sub cmd_Proces_Click()
   
   If cmb_Period.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Period)
      Exit Sub
   End If
   
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de generar el RCD?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   l_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00") & "01"
   l_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00") & Format(ff_Ultimo_Dia_Mes(cmb_Period.ItemData(cmb_Period.ListIndex), CInt(ipp_PerAno.Text)), "00")
            
   Screen.MousePointer = 11
   
      'Eliminamos el contenido de la tabla Identificacion si es q existiera
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "DELETE FROM TMP_DESRCD WHERE "
      g_str_Parame = g_str_Parame & "DESRCD_PERMES = " & Mid(l_str_FecIni, 5, 2) & " AND "
      g_str_Parame = g_str_Parame & "DESRCD_PERANO = " & Mid(l_str_FecIni, 1, 4) & " AND "
      g_str_Parame = g_str_Parame & "DESRCD_TERCRE ='" & modgen_g_str_NombPC & "' "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      'g_rst_Genera.Close
      'Set g_rst_Genera = Nothing
      
      'Eliminamos el contenido de la tabla Saldos si es q existiera
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "DELETE FROM TMP_SALRCD WHERE "
      g_str_Parame = g_str_Parame & "SALRCD_PERMES = " & Mid(l_str_FecIni, 5, 2) & " AND "
      g_str_Parame = g_str_Parame & "SALRCD_PERANO = " & Mid(l_str_FecIni, 1, 4) & " AND "
      g_str_Parame = g_str_Parame & "SALRCD_TERCRE ='" & modgen_g_str_NombPC & "' "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      'g_rst_Genera.Close
      'Set g_rst_Genera = Nothing
      
      Call fs_Genera_HipRcd(l_str_FecIni, l_str_FecFin, pnl_BarPro)
      Call fs_Genera_ComRcd(l_str_FecIni, l_str_FecFin, pnl_BarPro)
      
   Screen.MousePointer = 0
   
   MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
            
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
   
   cmb_Period.Clear

   Call moddat_gs_Carga_LisIte_Combo(cmb_Period, 1, "033")

   ipp_PerAno = Mid(date, 7, 4)
   
End Sub

Private Sub fs_Limpia()

   Dim r_int_PerMes  As Integer
   Dim r_int_PerAno  As Integer

   If Month(date) = 12 Then
      r_int_PerMes = 1
      r_int_PerAno = Year(date) - 1
   Else
      r_int_PerMes = Month(date) - 1
      r_int_PerAno = Year(date)
   End If

   Call gs_BuscarCombo_Item(cmb_Period, r_int_PerMes)
   ipp_PerAno.Text = Format(r_int_PerAno, "0000")
   
   pnl_BarPro.FloodPercent = 0
   
End Sub

Private Sub cmb_Period_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_PerAno)
    End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmd_Proces)
    End If
End Sub

Private Sub fs_Genera_HipRcd(ByVal p_FecIni As String, ByVal p_FecFin As String, Optional p_BarPro As SSPanel)
   
   Dim r_str_ApeCas        As String
   Dim r_str_CodSbs        As String
   Dim r_str_MagSbs        As String
   Dim r_str_MagEmp        As String
   Dim r_str_Sigla         As String
   Dim r_str_TipTri        As String
   Dim r_str_DocTri        As String
   Dim r_str_SegNom        As String
   Dim r_str_CodOfi        As String
   Dim r_str_CodSof        As String
   Dim r_str_TipIde        As String
   Dim r_str_NumTom        As String
   Dim r_str_NumPar        As String
   Dim r_str_NumFol        As String
   Dim r_str_Nombre        As String
   Dim r_str_PriNom        As String
   Dim r_str_LinGar        As String
      
   Dim r_int_NumSec        As Integer
   Dim r_int_CodCiu        As Integer
   Dim r_int_SalIte        As Integer
   Dim r_int_Contad        As Integer
      
   Dim r_dbl_TipCam        As Double
   
   Dim r_rst_SalRcd        As ADODB.Recordset
   Dim r_arr_CtaRcd()      As modtac_tpo_CtaRcd
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   r_str_Sigla = " "
   r_str_ApeCas = " "
      
   r_str_CodOfi = " "
   r_str_CodSof = " "
   r_str_TipIde = " "
   r_str_NumPar = " "
   r_str_NumFol = " "
   r_str_TipTri = 0
   r_str_DocTri = " "
   r_str_NumTom = " "
     
   r_str_MagEmp = 0
   r_int_NumSec = 1
   r_int_SalIte = 1
   
   
   r_lng_NumReg = 0
   r_lng_TotReg = ff_ConHip() + ff_ConCom()
   p_BarPro.FloodPercent = 0
   
   'Leyendo cursor Principal
   g_str_Parame = "SELECT HIPCIE_NUMOPE, HIPCIE_TDOCLI, HIPCIE_NDOCLI, HIPCIE_CLACLI, HIPCIE_CLAPRV, HIPCIE_CLAPRD, HIPCIE_ACUDIF, HIPCIE_SALCAP, HIPCIE_TIPMON, HIPCIE_ACUDVC, "
   g_str_Parame = g_str_Parame & "HIPCIE_MONGAR, HIPCIE_TIPGAR, HIPCIE_MTOGAR, DATGEN_TIPDOC, HIPCIE_CODPRD, HIPCIE_CAPVEN, HIPCIE_ACUDVG, HIPCIE_CAPVIG, HIPCIE_PRVESP, "
   g_str_Parame = g_str_Parame & "HIPCIE_SALCON, HIPCIE_DIAMOR, DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_NOMBRE, DATGEN_APECAS, DATGEN_CODSBS, DATGEN_OCUPAC, HIPCIE_PRVCAM, "
   g_str_Parame = g_str_Parame & "DATGEN_CODCIU, DATGEN_RESIDE, DATGEN_FLGACC, DATGEN_RELLAB, DATGEN_NACPAI, DATGEN_ESTCIV, DATGEN_CODSEX, DATGEN_UBIGEO, "
   g_str_Parame = g_str_Parame & "DATGEN_TDOTRI, DATGEN_NDOTRI, HIPCIE_TIPCAM FROM CRE_HIPCIE, CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_TDOCLI = DATGEN_TIPDOC AND "
   g_str_Parame = g_str_Parame & "HIPCIE_NDOCLI = DATGEN_NUMDOC AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & Mid(p_FecIni, 1, 4) & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & Mid(p_FecIni, 5, 2) & " "
   g_str_Parame = g_str_Parame & "ORDER BY DATGEN_APEPAT, DATGEN_APEMAT, DATGEN_APECAS, DATGEN_NOMBRE ASC "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_dbl_TipCam = g_rst_Princi!HIPCIE_TIPCAM
      
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
      g_rst_Princi.MoveFirst
   
      Do While Not g_rst_Princi.EOF
                  
         If r_int_SalIte <> 1 Then
            r_int_SalIte = 1
         End If
         
         'r_str_PriNom = " "
         'r_str_SegNom = " "
         r_str_Nombre = ff_Nombre(g_rst_Princi!DatGen_Nombre, r_str_PriNom, r_str_SegNom)
         r_str_LinGar = ff_LinGar(g_rst_Princi!HIPCIE_NUMOPE)
                  
         'Insertando Registro en Tabla de Identificación
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO TMP_DESRCD("
         g_str_Parame = g_str_Parame & "DESRCD_PERMES, "
         g_str_Parame = g_str_Parame & "DESRCD_PERANO, "
         g_str_Parame = g_str_Parame & "DESRCD_FERCRE, "
         g_str_Parame = g_str_Parame & "DESRCD_HORCRE, "
         g_str_Parame = g_str_Parame & "DESRCD_TERCRE, "
         g_str_Parame = g_str_Parame & "DESRCD_DESITE, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPFOR, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPINF, "
         g_str_Parame = g_str_Parame & "DESRCD_NUMSEC, "
         g_str_Parame = g_str_Parame & "DESRCD_CODSBS, "
         g_str_Parame = g_str_Parame & "DESRCD_CODINT, "
         g_str_Parame = g_str_Parame & "DESRCD_CODCIU, "
         g_str_Parame = g_str_Parame & "DESRCD_CODOFI, "
         g_str_Parame = g_str_Parame & "DESRCD_CODSOF, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPIDE, "
         g_str_Parame = g_str_Parame & "DESRCD_NUMPAR, "
         g_str_Parame = g_str_Parame & "DESRCD_NUMTOM, "
         g_str_Parame = g_str_Parame & "DESRCD_NUMFOL, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPTRI, "
         g_str_Parame = g_str_Parame & "DESRCD_DOCTRI, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPDOC, "
         g_str_Parame = g_str_Parame & "DESRCD_NUMDOC, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPPER, "
         g_str_Parame = g_str_Parame & "DESRCD_RESIDE, "
         g_str_Parame = g_str_Parame & "DESRCD_CLADEU, "
         g_str_Parame = g_str_Parame & "DESRCD_MAGSBS, "
         g_str_Parame = g_str_Parame & "DESRCD_ACCINF, "
         g_str_Parame = g_str_Parame & "DESRCD_RELLAB, "
         g_str_Parame = g_str_Parame & "DESRCD_PAIRES, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPGEN, "
         g_str_Parame = g_str_Parame & "DESRCD_ESTCIV, "
         g_str_Parame = g_str_Parame & "DESRCD_SIGLA,  "
         g_str_Parame = g_str_Parame & "DESRCD_APEPAT, "
         g_str_Parame = g_str_Parame & "DESRCD_APEMAT, "
         g_str_Parame = g_str_Parame & "DESRCD_APECAS, "
         g_str_Parame = g_str_Parame & "DESRCD_PRINOM, "
         g_str_Parame = g_str_Parame & "DESRCD_SEGNOM, "
         g_str_Parame = g_str_Parame & "DESRCD_RIECAM, "
         g_str_Parame = g_str_Parame & "DESRCD_INDATR, "
         g_str_Parame = g_str_Parame & "DESRCD_CLAREP) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         
         'Campos Basicos
         g_str_Parame = g_str_Parame & Mid(p_FecIni, 5, 2) & ", "
         g_str_Parame = g_str_Parame & Mid(p_FecIni, 1, 4) & ", "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         
         g_str_Parame = g_str_Parame & r_int_NumSec & ", "
         
         'Llenado de Datos para el RCD
         'Tipo de Formulario
         g_str_Parame = g_str_Parame & 1 & ", "
         
         'Tipo de Informacion
         g_str_Parame = g_str_Parame & 1 & ", "
         
         'Nro de Secuencia
         g_str_Parame = g_str_Parame & "'" & Format(CStr(r_int_NumSec), "00000000") & "', "
                  
         'Codigo Deudor SBS
         If IsNull(Trim(g_rst_Princi!DATGEN_CODSBS)) Then
            g_str_Parame = g_str_Parame & "'" & "0000000000" & "', "
         Else
            If Len(Trim(g_rst_Princi!DATGEN_CODSBS)) < 10 Then
               g_str_Parame = g_str_Parame & "'" & Format(Trim(g_rst_Princi!DATGEN_CODSBS), "0000000000") & "', "
            Else
               g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!DATGEN_CODSBS) & "', "
            End If
         End If
         
         'Codigo Deudor Asignado por la empresa Informante (TIPDOC + NUMDOC) (20 Caracteres Hacia la derecha en blanco)
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPCIE_TDOCLI & Trim(g_rst_Princi!HIPCIE_NDOCLI) & "', "
         
         'Codigo de CIUU si el deudor es dependiente se registra con 9999
         If g_rst_Princi!DATGEN_OCUPAC = "11" Then
            g_str_Parame = g_str_Parame & 9999 & ", "
         Else
            g_str_Parame = g_str_Parame & g_rst_Princi!DATGEN_CODCIU & ", "
         End If
                  
         'Codigo de Registro de Personas Juridicas
         'Para el Numero de Partida o Ficha Registral
         'A)Codigo de la Oficina Registral Regional
         
         g_str_Parame = g_str_Parame & "'" & r_str_CodOfi & "', "
         
         'B)Codigo de la Subsede de la Oficina Registral
         g_str_Parame = g_str_Parame & "'" & r_str_CodSof & "', "
         
         'c)Tipo de Informacion: P=Partida/F=Ficha | Tipo de Informacion: T=Tomo-Folio
         g_str_Parame = g_str_Parame & "'" & r_str_TipIde & "', "
         
         'D)Numero de Partida o Ficha
         g_str_Parame = g_str_Parame & "'" & r_str_NumPar & "', "
         
         'D1)Numero Tomo
         g_str_Parame = g_str_Parame & "'" & r_str_NumTom & "', "
         
         'E)Numero de Folio
         g_str_Parame = g_str_Parame & "'" & r_str_NumFol & "', "
         
         
         'Tipo de Documento Tributario
         If g_rst_Princi!DATGEN_TDOTRI <> 0 Then
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DATGEN_TDOTRI & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & r_str_TipTri & "', "
         End If
         
         'Documento Tributario
         If IsNull(g_rst_Princi!DATGEN_NDOTRI) Then
            g_str_Parame = g_str_Parame & "'" & r_str_DocTri & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!DATGEN_NDOTRI) & "', "
         End If
         
         'Tipo de Documento
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPCIE_TDOCLI & ", "
         
         'Nro de Documento
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPCIE_NDOCLI & "', "
         
         'Tipo de Persona 1-Persona Natural / 2-Persona Juridica / 3-Persona Mancomuna
         If r_str_DocTri = " " Then
            g_str_Parame = g_str_Parame & "'" & "1" & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & "2" & "', "
         End If
         
         'Residencia
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DATGEN_RESIDE & "', "
         
         'Clasificacion del Deudor 0-Normal / 1-CPP ...
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPCIE_CLACLI & "', "
         
         'Clasificacion de la Empresa 0-Persona Natural / 1-Persona Juridica Grande ...
         g_str_Parame = g_str_Parame & "'" & r_str_MagEmp & "', "
         
         'Accionista en la empresa Informante
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DATGEN_FLGACC & "', "
         
         'Relacion Laboral con la empresa Informante
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DATGEN_RELLAB & "', "
         
         'Pais de Residencia
         g_str_Parame = g_str_Parame & "'" & Format(Mid(g_rst_Princi!DATGEN_NACPAI, 3, 4), "0000") & "', "
         
         'Genero
         If g_rst_Princi!DatGen_CodSex = 1 Then
            g_str_Parame = g_str_Parame & "'" & "M" & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & "F" & "', "
         End If
      
         'Estado Civil
         If g_rst_Princi!DATGEN_ESTCIV = 1 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            g_str_Parame = g_str_Parame & "'" & "S" & "', "
         ElseIf g_rst_Princi!DATGEN_ESTCIV = 2 Then
            g_str_Parame = g_str_Parame & "'" & "C" & "', "
         ElseIf g_rst_Princi!DATGEN_ESTCIV = 3 Then
            g_str_Parame = g_str_Parame & "'" & "D" & "', "
         ElseIf g_rst_Princi!DATGEN_ESTCIV = 4 Then
            g_str_Parame = g_str_Parame & "'" & "V" & "', "
         End If
         
         'Sigla o Nombre Comercial
         g_str_Parame = g_str_Parame & "'" & r_str_Sigla & "', "
         
         'Apellido Paterno o Razon Social
         If IsNull(Trim(g_rst_Princi!DatGen_ApePat)) Then
            g_str_Parame = g_str_Parame & "'" & "XXXX" & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!DatGen_ApePat) & "', "
         End If
         
         'Apellido Materno
         If IsNull(Trim(g_rst_Princi!DatGen_ApeMat)) Then
            g_str_Parame = g_str_Parame & "'" & "XXXX" & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!DatGen_ApeMat) & "', "
         End If
         
         'Apellido de Casada
         If g_rst_Princi!DatGen_CodSex = 2 Then
            If (g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 4) Then
               If IsNull(Trim(g_rst_Princi!DatGen_ApeCas)) Then
                  g_str_Parame = g_str_Parame & "'" & r_str_ApeCas & "', "
               Else
                  g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!DatGen_ApeCas) & "', "
               End If
            Else
               g_str_Parame = g_str_Parame & "'" & r_str_ApeCas & "', "
            End If
         Else
            g_str_Parame = g_str_Parame & "'" & r_str_ApeCas & "', "
         End If
         
         'Primer Nombre
         g_str_Parame = g_str_Parame & "'" & Trim(r_str_PriNom) & "', "
                  
         'Segundo Nombre
         g_str_Parame = g_str_Parame & "'" & r_str_SegNom & "', "
         
         'Indicador de Riesgo Cambiario Crediticio
         If g_rst_Princi!HIPCIE_TIPMON = 2 Then
            g_str_Parame = g_str_Parame & "'" & 2 & "', "
         ElseIf g_rst_Princi!HIPCIE_TIPMON = 1 Then
            g_str_Parame = g_str_Parame & "'" & 1 & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & 0 & "', "
         End If
         
         'Indicador de Atraso del Deudor
         g_str_Parame = g_str_Parame & "'" & "A" & "', "
         
         'Clasificacion del deudor de la empresa reportante
         g_str_Parame = g_str_Parame & "'" & "XXXXX" & "') "
                             
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
                  
         'Para leer cuentas para RCD
         ReDim r_arr_CtaRcd(0)
         
         'Leyendo Tabla de Cuentas y llenado en Arreglo
         g_str_Parame = "SELECT * FROM TMP_CUERCD WHERE "
         g_str_Parame = g_str_Parame & "CUERCD_PERANO =" & Mid(p_FecIni, 1, 4) & " AND "
         g_str_Parame = g_str_Parame & "CUERCD_PERMES =" & Mid(p_FecIni, 5, 2) & " "
         g_str_Parame = g_str_Parame & "ORDER BY CUERCD_CTACTB ASC "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_SalRcd, 3) Then
            Exit Sub
         End If
         
         If Not (r_rst_SalRcd.BOF And r_rst_SalRcd.EOF) Then
         
            r_rst_SalRcd.MoveFirst
         
            Do While Not r_rst_SalRcd.EOF
            
               ReDim Preserve r_arr_CtaRcd(UBound(r_arr_CtaRcd) + 1)
               
               r_arr_CtaRcd(UBound(r_arr_CtaRcd)).CtaRcd_NumCta = Trim(r_rst_SalRcd!CUERCD_CTACTB)
               r_arr_CtaRcd(UBound(r_arr_CtaRcd)).CtaRcd_DesVar = Trim(r_rst_SalRcd!CUERCD_DESVAR)
               r_arr_CtaRcd(UBound(r_arr_CtaRcd)).CtaRcd_Import = 0
               
               r_rst_SalRcd.MoveNext
            
            Loop
            
            r_rst_SalRcd.Close
            Set r_rst_SalRcd = Nothing
            
         End If
                  
         For r_int_Contad = 1 To UBound(r_arr_CtaRcd)
                                                
            Select Case r_arr_CtaRcd(r_int_Contad).CtaRcd_DesVar
                              
               'Interes Diferido en MN - 29110201000000
               Case "acdfmn":
                              If g_rst_Princi!HIPCIE_ACUDIF > 0 Then
                                 If g_rst_Princi!HIPCIE_TIPMON = 1 Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_ACUDIF
                                 End If
                              End If
                              
               'Interes Diferido en ME - 29210201000000
               Case "acdfme":
                              If g_rst_Princi!HIPCIE_ACUDIF > 0 Then
                                 If g_rst_Princi!HIPCIE_TIPMON = 2 Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!HIPCIE_ACUDIF * r_dbl_TipCam, "###,###,##0.00")
                                 End If
                              End If
               'Hipotecas Preferidas en MN - 84140201000000
               Case "hpprmn":
                              If g_rst_Princi!HIPCIE_MONGAR = 1 Then
                                 If ((g_rst_Princi!HIPCIE_TIPGAR = 1) Or (g_rst_Princi!HIPCIE_TIPGAR = 2)) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_MTOGAR
                                 End If
                              End If
                              
               'Hipotecas Preferidas en ME - 84240201000000
               Case "hpprme":
                              If g_rst_Princi!HIPCIE_MONGAR = 2 Then
                                 If ((g_rst_Princi!HIPCIE_TIPGAR = 1) Or (g_rst_Princi!HIPCIE_TIPGAR = 2)) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!HIPCIE_MTOGAR * r_dbl_TipCam, "###,###,##0.00")
                                 End If
                              End If
               
               'Hipotecas Fianza Solidaria en MN - 84191901010100
               'Fianza solidaria se toma los saldos TNC + TC
               'Case "fisomn":
               '               If g_rst_Princi!HIPCIE_TIPMON = 1 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 3 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON
               '                  End If
               '               End If
               
               'Hipotecas Carta Fianza en MN - 84191901010200
               'Carta Fianza se toma los saldos TNC + TC
               'Case "cafimn":
               '               If g_rst_Princi!HIPCIE_MONGAR = 1 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 4 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_MTOGAR
               '                  End If
               '               End If
                              
               'Hipotecas Retencion de Fondos en MN - 84140900000000
               'Retencion de Fondos se toma los saldos TNC + TC
               'Case "refomn":
               '               If g_rst_Princi!HIPCIE_TIPMON = 1 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 6 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON
               '                  End If
               '               End If
               
               'Hipotecas Certificado de Participacion en MN - 84140900000000
               'Certificado de Participacion se toma los saldos TNC + TC
               'Case "cepamn":
               '               If g_rst_Princi!HIPCIE_TIPMON = 1 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 5 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON
               '                  End If
               '               End If
               
               'Hipotecas Fianza Solidaria en MN - 84291901010100
               'Fianza solidaria se toma los saldos TNC + TC
               'Case "fisome":
               '               If g_rst_Princi!HIPCIE_TIPMON = 2 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 3 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON
               '                  End If
               '               End If
               
               'Hipotecas Carta Fianza en MN - 84291901010200
               'Carta Fianza se toma los saldos TNC + TC
               'Case "cafime":
               '               If g_rst_Princi!HIPCIE_MONGAR = 2 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 4 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_MTOGAR
               '                  End If
               '               End If
                              
               'Hipotecas Retencion de Fondos en MN - 84240900000000
               'Retencion de Fondos se toma los saldos TNC + TC
               'Case "refome":
               '               If g_rst_Princi!HIPCIE_TIPMON = 2 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 6 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON
               '                  End If
               '               End If
               
               
               'Hipotecas Certificado de Participacion en MN - 84240900000000
               'Certificado de Participacion se toma los saldos TNC + TC
               'Case "cepame":
               '               If g_rst_Princi!HIPCIE_TIPMON = 2 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 5 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON
               '                  End If
               '               End If
               
               'Hipotecas Fianza Solidaria/Carta Fianza en MN - 84140501000000
               'Carta Fianza se toma la garantia y Fianza solidaria se toma los saldos TNC + TC
               'Case "crfimn":
               '               If g_rst_Princi!HIPCIE_TIPMON = 1 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 3 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON
               '                  End If
               '               End If
               '
               '               If g_rst_Princi!HIPCIE_MONGAR = 1 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 4 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_MTOGAR
               '                  End If
               '               End If
                              
               'Hipotecas Fianza Solidaria/Carta Fianza en ME - 84240501000000
               'Carta Fianza se toma la garantia y Fianza solidaria se toma los saldos TNC + TC
               'Case "fisome":
               '               If g_rst_Princi!HIPCIE_TIPMON = 2 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 3 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * r_dbl_TipCam, "###,###,##0.00")
               '                  End If
               '               End If
               '
               '               If g_rst_Princi!HIPCIE_MONGAR = 2 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 4 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!HIPCIE_MTOGAR * r_dbl_TipCam, "###,###,##0.00")
               '                  End If
               '               End If
               
               'Hipotecas Certificado de Participacion/Retencion de Fondos en MN - 84140900000000
               'Certificado de Participacion se toma la garantia y Retencion de Fondos se toma los saldos TNC + TC
               'Case "grprmn":
               '               If g_rst_Princi!HIPCIE_TIPMON = 1 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 6 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON
               '                  End If
               '               End If
               '
               '               If g_rst_Princi!HIPCIE_MONGAR = 1 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 5 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_MTOGAR
               '                  End If
               '               End If
                              
               'Hipotecas Certificado de Participacion/Retencion de Fondos en ME - 84240900000000
               'Certificado de Participacion se toma la garantia y Retencion de Fondos se toma los saldos TNC + TC
               'Case "refome":
               '               If g_rst_Princi!HIPCIE_TIPMON = 2 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 6 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * r_dbl_TipCam, "###,###,##0.00")
               '                  End If
               '               End If
               '
               '               If g_rst_Princi!HIPCIE_MONGAR = 2 Then
               '                  If g_rst_Princi!HIPCIE_TIPGAR = 5 Then
               '                     r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!HIPCIE_MTOGAR * r_dbl_TipCam, "###,###,##0.00")
               '                  End If
               '               End If
               
               
               'Creditos Vencidos en Suspenso en MN - 81140200000000
               Case "hpvnmn":
                              If g_rst_Princi!HIPCIE_ACUDVC > 0 Then
                                 If g_rst_Princi!HIPCIE_TIPMON = 1 Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_ACUDVC
                                 End If
                              End If
               'Creditos Vencidos en Suspenso en ME - 81240200000000
               Case "hpvnme":
                              If g_rst_Princi!HIPCIE_ACUDVC > 0 Then
                                 If g_rst_Princi!HIPCIE_TIPMON = 2 Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format((g_rst_Princi!HIPCIE_ACUDVC * r_dbl_TipCam), "###,###,##0.00")
                                 End If
                              End If
                              
               'Prestamos del Fondo MI-VIVIENDA miHogar y miVivienda en MN - 14110423000000
               Case "prmvmn":
                              If (g_rst_Princi!HIPCIE_CODPRD = "004" Or g_rst_Princi!HIPCIE_CODPRD = "007") Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 1) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_CAPVIG
                                 End If
                              End If
                              
               'Prestamos del Fondo MI-VIVIENDA CME en MN - 14110425000000
               Case "prvemn":
                              If (g_rst_Princi!HIPCIE_CODPRD = "003") Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 1) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_CAPVIG
                                 End If
                              End If
                              
               'Prestamos del Fondo MI-VIVIENDA con Recursos de Instituciones Financieras CRC-PBP ME - 14210424000000
               Case "prmvme":
                              If (g_rst_Princi!HIPCIE_CODPRD = "001") Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 2) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format((g_rst_Princi!HIPCIE_CAPVIG) * r_dbl_TipCam, "###,###,##0.00")
                                 End If
                              End If
               
               'Prestamos MiCasita Capital Vencido ME - 14250406000000
               Case "cpmcme":
                              If (g_rst_Princi!HIPCIE_CAPVEN > 0) Then
                                 If (g_rst_Princi!HIPCIE_CODPRD = "002") Then
                                    If (g_rst_Princi!HIPCIE_TIPMON = 2) Then
                                       r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!HIPCIE_CAPVEN * r_dbl_TipCam, "###,###,##0.00")
                                    End If
                                 End If
                              End If
                              
               'Prestamos Fondo MiVivienda Capital Vencido MN - 14150423000000
               Case "cpmhmn":
                              If (g_rst_Princi!HIPCIE_CAPVEN > 0) Then
                                 If (g_rst_Princi!HIPCIE_CODPRD = "004" Or g_rst_Princi!HIPCIE_CODPRD = "007") Then
                                    If (g_rst_Princi!HIPCIE_TIPMON = 1) Then
                                       r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_CAPVEN
                                    End If
                                 End If
                              End If
               
               'Prestamos Fondo MiVivienda CRC-PBP Capital Vencido ME - 14250424000000
               Case "pcrcvn":
                              If (g_rst_Princi!HIPCIE_CAPVEN > 0) Then
                                 If (g_rst_Princi!HIPCIE_CODPRD = "001") Then
                                    If (g_rst_Princi!HIPCIE_TIPMON = 2) Then
                                       r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!HIPCIE_CAPVEN * r_dbl_TipCam, "###,###,##0.00")
                                    End If
                                 End If
                              End If
                              
               'Prestamos Fondo MiVivienda CME Capital Vencido ME - 14150425000000
               Case "cpcemn":
                              If (g_rst_Princi!HIPCIE_CAPVEN > 0) Then
                                 If (g_rst_Princi!HIPCIE_CODPRD = "003") Then
                                    If (g_rst_Princi!HIPCIE_TIPMON = 1) Then
                                       r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_CAPVEN
                                    End If
                                 End If
                              End If
                              
               'Interes Devengado Acumulado de Creditos Hipotecarios MN - 14180400000000
               Case "dvhpmn":
                              If (g_rst_Princi!HIPCIE_ACUDVG > 0) Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 1) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_ACUDVG
                                 End If
                              End If
               
               'Interes Devengado Acumulado de Creditos Hipotecarios ME - 14280400000000
               Case "dvhpme":
                              If (g_rst_Princi!HIPCIE_ACUDVG > 0) Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 2) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!HIPCIE_ACUDVG * r_dbl_TipCam, "###,###,##0.00")
                                 End If
                              End If
                              
               'Prestamos miCasita con Hipoteca Inscrita ME - 14210406010000
               Case "hpvgme":
                              If (g_rst_Princi!HIPCIE_CODPRD = "002") Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 2) Then
                                    If (g_rst_Princi!HIPCIE_TIPGAR = 1 Or g_rst_Princi!HIPCIE_TIPGAR = 2) Then
                                       r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!HIPCIE_CAPVIG * r_dbl_TipCam, "###,###,##0.00")
                                    End If
                                 End If
                              End If
                              
               'Prestamos miCasita con Hipoteca Inscrita MN - 14110406010000
               Case "hpvgmn":
                              If (g_rst_Princi!HIPCIE_CODPRD = "006") Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 1) Then
                                    If (g_rst_Princi!HIPCIE_TIPGAR = 1 Or g_rst_Princi!HIPCIE_TIPGAR = 2) Then
                                       r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON
                                    End If
                                 End If
                              End If
               
               'Prestamos miCasita sin Hipoteca Inscrita ME - 14210406020000
               Case "snhpme":
                              If (g_rst_Princi!HIPCIE_CODPRD = "002") Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 2) Then
                                    If (g_rst_Princi!HIPCIE_TIPGAR > 2) Then
                                       r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format((g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON) * r_dbl_TipCam, "###,###,##0.00")
                                    End If
                                 End If
                              End If
                              
               'Prestamos miCasita sin Hipoteca Inscrita MN - 14110406020000
               Case "prshmn":
                              If (g_rst_Princi!HIPCIE_CODPRD = "006") Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 1) Then
                                    If (g_rst_Princi!HIPCIE_TIPGAR > 2) Then
                                       r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON
                                    End If
                                 End If
                              End If
               
               'Provisiones para Creditos Hipotecarios Especificas en MN - 14290401000000
               Case "prhimn":
                              If (g_rst_Princi!HIPCIE_PRVESP > 0) Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 1) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_PRVESP
                                 End If
                              End If
                              
               'Provisiones para Creditos Hipotecarios Especificas en ME - 14190401000000
               Case "prhime":
                              If (g_rst_Princi!HIPCIE_PRVESP > 0) Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 2) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_PRVESP
                                 End If
                              End If
                              
               'Provisiones para Creditos Hipotecarios Riesgo Cambiario Crediticio en ME - 14290405000000
               Case "prhpme":
                              If (g_rst_Princi!HIPCIE_PRVCAM > 0) Then
                                 If (g_rst_Princi!HIPCIE_TIPMON = 2) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!HIPCIE_PRVCAM
                                 End If
                              End If
                          
            End Select
            
            If r_arr_CtaRcd(r_int_Contad).CtaRcd_Import <> 0 Then
               
               'Insertando Registro de Saldos por Cliente
               g_str_Parame = ""
               g_str_Parame = g_str_Parame & "INSERT INTO TMP_SALRCD("
               g_str_Parame = g_str_Parame & "SALRCD_PERMES, "
               g_str_Parame = g_str_Parame & "SALRCD_PERANO, "
               g_str_Parame = g_str_Parame & "SALRCD_FERCRE, "
               g_str_Parame = g_str_Parame & "SALRCD_HORCRE, "
               g_str_Parame = g_str_Parame & "SALRCD_TERCRE, "
               g_str_Parame = g_str_Parame & "SALRCD_SALITE, "
               g_str_Parame = g_str_Parame & "SALRCD_TIPFOR, "
               g_str_Parame = g_str_Parame & "SALRCD_TIPINF, "
               g_str_Parame = g_str_Parame & "SALRCD_NUMSEC, "
               g_str_Parame = g_str_Parame & "SALRCD_TIPDOC, "
               g_str_Parame = g_str_Parame & "SALRCD_NUMDOC, "
               g_str_Parame = g_str_Parame & "SALRCD_CODOFI, "
               g_str_Parame = g_str_Parame & "SALRCD_UBIGEO, "
               g_str_Parame = g_str_Parame & "SALRCD_CTACTB, "
               g_str_Parame = g_str_Parame & "SALRCD_TIPCRE, "
               g_str_Parame = g_str_Parame & "SALRCD_MTOSDO, "
               g_str_Parame = g_str_Parame & "SALRCD_CONDIA, "
               g_str_Parame = g_str_Parame & "SALRCD_CONCTA, "
               g_str_Parame = g_str_Parame & "SALRCD_CONDIS) "
               
               
               g_str_Parame = g_str_Parame & "VALUES ("
               
               'Datos Basicos
               g_str_Parame = g_str_Parame & Mid(p_FecIni, 5, 2) & ", "
               g_str_Parame = g_str_Parame & Mid(p_FecIni, 1, 4) & ", "
               g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
               g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
               
               'Nro de Iteracion
               g_str_Parame = g_str_Parame & r_int_SalIte & ", "
               
               'Tipo de Formulario 1 = Del Deudor / 2 = Totales de la Empresa
               g_str_Parame = g_str_Parame & 1 & ", "
               
               'Tipo de Informacion
               g_str_Parame = g_str_Parame & 2 & ", "
               
               'Nro de Secuencia
               g_str_Parame = g_str_Parame & "'" & Format(CStr(r_int_NumSec), "00000000") & "', "
               
               'Tipo de Documento
               g_str_Parame = g_str_Parame & g_rst_Princi!HIPCIE_TDOCLI & ", "
               
               'Nro de Documento
               g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!HIPCIE_NDOCLI) & "', "
               
               'Codigo de la Empresa Informante
               g_str_Parame = g_str_Parame & "'" & Format("0001", "0000") & "', "
               
               'Ubicacion Geografica de la Oficina de la empresa Informante (Dpto-Prov-Dist)
               g_str_Parame = g_str_Parame & "'" & Format("150131", "000000") & "', "
               
               'Codigo de Cuenta Contable
               g_str_Parame = g_str_Parame & "'" & Format(r_arr_CtaRcd(r_int_Contad).CtaRcd_NumCta, "00000000000000") & "', "
               
               'Tipo de Credito
               g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPCIE_CLAPRD & "', "
               
               'Saldo
               g_str_Parame = g_str_Parame & r_arr_CtaRcd(r_int_Contad).CtaRcd_Import & ", "
               
               'Condicion en dias
               g_str_Parame = g_str_Parame & "'" & Format(Trim(g_rst_Princi!HIPCIE_DIAMOR), "0000") & "', "
               
               'Condicion de disponibilidad / Linea de Garantia
               If r_str_LinGar = "999999" Then
                  g_str_Parame = g_str_Parame & "'" & "02" & "', "
               Else
                  g_str_Parame = g_str_Parame & "'" & "01" & "', "
               End If
               
               'Condicion especial de la cuenta / Tipo de Garantia
               If g_rst_Princi!HIPCIE_TIPGAR = "1" Then
                  g_str_Parame = g_str_Parame & "'" & "04" & "') "
               ElseIf g_rst_Princi!HIPCIE_TIPGAR = "2" Then
                  g_str_Parame = g_str_Parame & "'" & "05" & "') "
               Else
                  g_str_Parame = g_str_Parame & "'" & "06" & "') "
               End If
                          
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                  Exit Sub
               End If
               
               r_int_SalIte = r_int_SalIte + 1
               
            End If
            
         Next r_int_Contad
                          
         r_int_NumSec = r_int_NumSec + 1
         
         r_lng_NumReg = r_lng_NumReg + 1
                  
         g_rst_Princi.MoveNext
         DoEvents
         
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
   Else
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
 
      Screen.MousePointer = 0
 
      MsgBox "No se encontraron Datos registradas.", vbInformation, modgen_g_str_NomPlt
 
      Exit Sub
   End If
     

End Sub

Private Sub fs_Genera_ComRcd(ByVal p_FecIni As String, ByVal p_FecFin As String, Optional p_BarPro As SSPanel)
   
   Dim r_str_ApeCas        As String
   Dim r_str_CodSbs        As String
   Dim r_str_MagSbs        As String
   Dim r_str_MagEmp        As String
   Dim r_str_Sigla         As String
   Dim r_str_TipTri        As String
   Dim r_str_DocTri        As String
   Dim r_str_SegNom        As String
   Dim r_str_CodOfi        As String
   Dim r_str_CodSof        As String
   Dim r_str_TipIde        As String
   Dim r_str_NumPar        As String
   Dim r_str_NumFol        As String
   Dim r_str_ComCie        As String
   Dim r_str_ClaCli        As String
   Dim r_str_ClaPrv        As String
   Dim r_str_ClaPrd        As String
   Dim r_str_DiaMor        As String
   Dim r_str_NumTom        As String
      
   Dim r_int_NumSec        As Integer
   Dim r_int_CodCiu        As Integer
   Dim r_int_SalIte        As Integer
   Dim r_int_Contad        As Integer
   
   Dim r_dbl_TipCam        As Double
   
   Dim r_rst_SalRcd        As ADODB.Recordset
   Dim r_arr_CtaRcd()      As modtac_tpo_CtaRcd
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
   
   r_str_Sigla = " "
   r_str_ApeCas = " "
   
   r_str_CodOfi = " "
   r_str_CodSof = " "
   r_str_TipIde = " "
   r_str_NumPar = " "
   r_str_NumFol = " "
   r_str_TipTri = " "
   r_str_DocTri = " "
   r_str_NumTom = " "
     
   r_str_MagEmp = 0
   r_int_NumSec = 1
   r_int_SalIte = 1
      
   
   'Leyendo cursor Principal
   g_str_Parame = "SELECT DISTINCT(COMCIE_NDOCLI) AS NDOCLI, MAX(COMCIE_TDOCLI) AS TDOCLI, MAX(COMCIE_CLACLI) AS CLACLI, MAX(COMCIE_CLAPRV) AS CLAPRV, MAX(COMCIE_TIPMON) AS TIPMON, SUM(COMCIE_ACUDVG) AS ACUDVG, "
   g_str_Parame = g_str_Parame & "MAX(COMCIE_CLAPRD) AS CLAPRD, MAX(COMCIE_DIAMOR) AS DIAMOR, MAX(COMCIE_TIPCAM) AS TIPCAM, SUM(COMCIE_MTOGAR) AS MTOGAR, "
   g_str_Parame = g_str_Parame & "SUM(COMCIE_SALCAP) AS SALCAP, SUM(COMCIE_PRVCAM) AS PRVCAM, MAX(COMCIE_MONGAR) AS MONGAR, MAX(DATGEN_PAIRES) AS PAIRES, MAX(DATGEN_FLGACC) AS FLGACC,"
   g_str_Parame = g_str_Parame & "MAX(DATGEN_CODSBS) AS CODSBS, MAX(DATGEN_MAGSBS) AS MAGSBS, MAX(DATGEN_CODCIU) AS CODCIU, MAX(COMCIE_TIPGAR) AS TIPGAR,  "
   g_str_Parame = g_str_Parame & "MAX(DATGEN_NOMCOM) AS NOMCOM, MAX(DATGEN_RAZSOC) AS RAZSOC, MAX(DATGEN_UBIGEO) AS UBIGEO FROM CRE_COMCIE, EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_TDOCLI = DATGEN_EMPTDO AND "
   g_str_Parame = g_str_Parame & "COMCIE_NDOCLI = DATGEN_EMPNDO AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & Mid(p_FecIni, 1, 4) & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & Mid(p_FecIni, 5, 2) & " "
   g_str_Parame = g_str_Parame & "GROUP BY COMCIE_NDOCLI "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_dbl_TipCam = g_rst_Princi!TIPCAM
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      
      g_rst_Princi.MoveFirst
   
      Do While Not g_rst_Princi.EOF
                  
         If r_int_SalIte <> 1 Then
            r_int_SalIte = 1
         End If
         
         'r_str_ComCie = ff_ComCie(p_FecIni, g_rst_Princi!DATGEN_EMPNDO, r_str_ClaCli, r_str_ClaPrv, r_str_ClaPrd, r_str_DiaMor)
                  
         'Insertando Registro en Tabla de Identificación
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO TMP_DESRCD("
         g_str_Parame = g_str_Parame & "DESRCD_PERMES, "
         g_str_Parame = g_str_Parame & "DESRCD_PERANO, "
         g_str_Parame = g_str_Parame & "DESRCD_FERCRE, "
         g_str_Parame = g_str_Parame & "DESRCD_HORCRE, "
         g_str_Parame = g_str_Parame & "DESRCD_TERCRE, "
         g_str_Parame = g_str_Parame & "DESRCD_DESITE, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPFOR, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPINF, "
         g_str_Parame = g_str_Parame & "DESRCD_NUMSEC, "
         g_str_Parame = g_str_Parame & "DESRCD_CODSBS, "
         g_str_Parame = g_str_Parame & "DESRCD_CODINT, "
         g_str_Parame = g_str_Parame & "DESRCD_CODCIU, "
         g_str_Parame = g_str_Parame & "DESRCD_CODOFI, "
         g_str_Parame = g_str_Parame & "DESRCD_CODSOF, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPIDE, "
         g_str_Parame = g_str_Parame & "DESRCD_NUMPAR, "
         g_str_Parame = g_str_Parame & "DESRCD_NUMTOM, "
         g_str_Parame = g_str_Parame & "DESRCD_NUMFOL, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPTRI, "
         g_str_Parame = g_str_Parame & "DESRCD_DOCTRI, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPDOC, "
         g_str_Parame = g_str_Parame & "DESRCD_NUMDOC, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPPER, "
         g_str_Parame = g_str_Parame & "DESRCD_RESIDE, "
         g_str_Parame = g_str_Parame & "DESRCD_CLADEU, "
         g_str_Parame = g_str_Parame & "DESRCD_MAGSBS, "
         g_str_Parame = g_str_Parame & "DESRCD_ACCINF, "
         g_str_Parame = g_str_Parame & "DESRCD_RELLAB, "
         g_str_Parame = g_str_Parame & "DESRCD_PAIRES, "
         g_str_Parame = g_str_Parame & "DESRCD_TIPGEN, "
         g_str_Parame = g_str_Parame & "DESRCD_ESTCIV, "
         g_str_Parame = g_str_Parame & "DESRCD_SIGLA,  "
         g_str_Parame = g_str_Parame & "DESRCD_APEPAT, "
         g_str_Parame = g_str_Parame & "DESRCD_APEMAT, "
         g_str_Parame = g_str_Parame & "DESRCD_APECAS, "
         g_str_Parame = g_str_Parame & "DESRCD_PRINOM, "
         g_str_Parame = g_str_Parame & "DESRCD_SEGNOM, "
         g_str_Parame = g_str_Parame & "DESRCD_RIECAM, "
         g_str_Parame = g_str_Parame & "DESRCD_INDATR, "
         g_str_Parame = g_str_Parame & "DESRCD_CLAREP) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         
         'Campos Basicos
         g_str_Parame = g_str_Parame & Mid(p_FecIni, 5, 2) & ", "
         g_str_Parame = g_str_Parame & Mid(p_FecIni, 1, 4) & ", "
         g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         
         g_str_Parame = g_str_Parame & r_int_NumSec & ", "
         
         'Llenado de Datos para el RCD
         'Tipo de Formulario
         g_str_Parame = g_str_Parame & 1 & ", "
         
         'Tipo de Informacion
         g_str_Parame = g_str_Parame & 1 & ", "
         
         'Nro de Secuencia
         g_str_Parame = g_str_Parame & "'" & Format(CStr(r_int_NumSec), "00000000") & "', "
                  
         'Codigo Deudor SBS
         If IsNull(Trim(g_rst_Princi!CODSBS)) Then
            g_str_Parame = g_str_Parame & "'" & "0000000000" & "', "
         Else
            If Len(Trim(g_rst_Princi!CODSBS)) < 10 Then
               g_str_Parame = g_str_Parame & "'" & Format(Trim(g_rst_Princi!CODSBS), "0000000000") & "', "
            Else
               g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!CODSBS) & "', "
            End If
         End If
         
         'Codigo Deudor Asignado por la empresa Informante (TIPDOC + NUMDOC) (20 Caracteres Hacia la derecha en blanco)
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!NDOCLI) & "', "
         
         'Codigo de CIUU si el deudor es dependiente se registra con 9999
         g_str_Parame = g_str_Parame & g_rst_Princi!CODCIU & ", "
                           
         'Codigo de Registro de Personas Juridicas
         'Para el Numero de Partida o Ficha Registral
         'A)Codigo de la Oficina Registral Regional
         g_str_Parame = g_str_Parame & "'" & r_str_CodOfi & "', "
         
         'B)Codigo de la Subsede de la Oficina Registral
         g_str_Parame = g_str_Parame & "'" & r_str_CodSof & "', "
         
         'c)Tipo de Informacion: P=Partida/F=Ficha | Tipo de Informacion: T=Tomo-Folio
         g_str_Parame = g_str_Parame & "'" & r_str_TipIde & "', "
         
         'D)Numero de Partida o Ficha
         g_str_Parame = g_str_Parame & "'" & r_str_NumPar & "', "
         
         'D1)Numero Tomo
         g_str_Parame = g_str_Parame & "'" & r_str_NumTom & "', "
         
         'E)Numero de Folio
         g_str_Parame = g_str_Parame & "'" & r_str_NumFol & "', "
                           
         'Tipo de Documento Tributario
         If Len(Trim(g_rst_Princi!NDOCLI)) = 11 Then
            g_str_Parame = g_str_Parame & "'" & 3 & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & 2 & "', "
         End If
         'Documento Tributario
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!NDOCLI) & "', "
         
         'Tipo de Documento
         g_str_Parame = g_str_Parame & 0 & ", "
         
         'Nro de Documento
         g_str_Parame = g_str_Parame & "'" & " " & "', "
         
         'Tipo de Persona 1-Persona Natural / 2-Persona Juridica / 3-Persona Mancomuna
         If IsNull(Trim(g_rst_Princi!NDOCLI)) Then
            g_str_Parame = g_str_Parame & "'" & "1" & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & "2" & "', "
         End If
         
         'Residencia
         g_str_Parame = g_str_Parame & "'" & 1 & "', "
         
         'Clasificacion del Deudor 0-Normal / 1-CPP ...
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CLACLI & "', "
         
         'Clasificacion de la Empresa 0-Persona Natural / 1-Persona Juridica Grande ...
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!MAGSBS & "', "
         
         'Accionista en la empresa Informante
         If IsNull(g_rst_Princi!FLGACC) Then
            g_str_Parame = g_str_Parame & "'" & 0 & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!FLGACC & "', "
         End If
         
         'Relacion Laboral con la empresa Informante
         g_str_Parame = g_str_Parame & "'" & 0 & "', "
         
         'Pais de Residencia
         g_str_Parame = g_str_Parame & "'" & Format(4028, "0000") & "', "
         
         'Genero
         g_str_Parame = g_str_Parame & "'" & 0 & "', "
                  
         'Estado Civil
         g_str_Parame = g_str_Parame & "'" & 0 & "', "
         
         'Sigla o Nombre Comercial
         g_str_Parame = g_str_Parame & "'" & Left(Trim(g_rst_Princi!NOMCOM), 20) & "', "
         
         'Apellido Paterno o Razon Social
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!RAZSOC) & "', "
         
         'Apellido Materno
         g_str_Parame = g_str_Parame & "'" & " " & "', "
         
         'Apellido de Casada
         g_str_Parame = g_str_Parame & "'" & " " & "', "
         
         'Primer Nombre
         g_str_Parame = g_str_Parame & "'" & " " & "', "
                  
         'Segundo Nombre
         g_str_Parame = g_str_Parame & "'" & " " & "', "
         
         'Indicador de Riesgo Cambiario Crediticio
         If g_rst_Princi!TIPMON = 2 Then
            g_str_Parame = g_str_Parame & "'" & 2 & "', "
         ElseIf g_rst_Princi!TIPMON = 1 Then
            g_str_Parame = g_str_Parame & "'" & 1 & "', "
         Else
            g_str_Parame = g_str_Parame & "'" & 0 & "', "
         End If
         'Indicador de Atraso del Deudor
         g_str_Parame = g_str_Parame & "'" & "A" & "', "
         
         'Clasificacion del deudor de la empresa reportante
         g_str_Parame = g_str_Parame & "'" & "XXXXX" & "') "
                             
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
                  
         'Para leer cuentas para RCD
         ReDim r_arr_CtaRcd(0)
         
         'Leyendo Tabla de Cuentas y llenado en Arreglo
         g_str_Parame = "SELECT * FROM TMP_CUERCD WHERE "
         g_str_Parame = g_str_Parame & "CUERCD_PERANO =" & Mid(p_FecIni, 1, 4) & " AND "
         g_str_Parame = g_str_Parame & "CUERCD_PERMES =" & Mid(p_FecIni, 5, 2) & " "
         g_str_Parame = g_str_Parame & "ORDER BY CUERCD_CTACTB ASC "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_SalRcd, 3) Then
            Exit Sub
         End If
         
         If Not (r_rst_SalRcd.BOF And r_rst_SalRcd.EOF) Then
         
            r_rst_SalRcd.MoveFirst
         
            Do While Not r_rst_SalRcd.EOF
            
               ReDim Preserve r_arr_CtaRcd(UBound(r_arr_CtaRcd) + 1)
               
               r_arr_CtaRcd(UBound(r_arr_CtaRcd)).CtaRcd_NumCta = Trim(r_rst_SalRcd!CUERCD_CTACTB)
               r_arr_CtaRcd(UBound(r_arr_CtaRcd)).CtaRcd_DesVar = Trim(r_rst_SalRcd!CUERCD_DESVAR)
               r_arr_CtaRcd(UBound(r_arr_CtaRcd)).CtaRcd_Import = 0
               
               r_rst_SalRcd.MoveNext
            
            Loop
            
            r_rst_SalRcd.Close
            Set r_rst_SalRcd = Nothing
            
         End If
                  
         For r_int_Contad = 1 To UBound(r_arr_CtaRcd)
                                                
            Select Case r_arr_CtaRcd(r_int_Contad).CtaRcd_DesVar
                              
               'Hipotecas Preferidas en MN - 84140201000000
               Case "hpprmn":
                              If g_rst_Princi!MONGAR = 1 Then
                                 If ((g_rst_Princi!TIPGAR = 1) Or (g_rst_Princi!TIPGAR = 2)) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!MTOGAR
                                 End If
                              End If
               
               'Hipotecas Preferidas en ME - 84240201000000
               Case "hpprme":
                              If g_rst_Princi!MONGAR = 2 Then
                                 If ((g_rst_Princi!TIPGAR = 1) Or (g_rst_Princi!TIPGAR = 2)) Then
                                    r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!MTOGAR * r_dbl_TipCam, "###,###,##0.00")
                                 End If
                              End If
               'Creditos Inmobiliarios en MN -14110127000000
               Case "crinmn":
                              If g_rst_Princi!TIPMON = 1 Then
                                 r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!SalCap
                              End If
               'Creditos Inmobiliarios en ME -14210127000000
               Case "crinme":
                              If g_rst_Princi!TIPMON = 2 Then
                                 r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!SalCap * r_dbl_TipCam, "###,###,##0.00")
                              End If
                              
               'Interes Devengado Acumulado de Creditos Inmobiliarios MN - 14180100000000
               Case "dvinmn":
                              If (g_rst_Princi!TIPMON = 1) Then
                                 r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = g_rst_Princi!ACUDVG
                              End If
               
               'Interes Devengado Acumulado de Creditos Inmobiliarios ME - 14280100000000
               Case "dvinme":
                              If (g_rst_Princi!TIPMON = 2) Then
                                 r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!ACUDVG * r_dbl_TipCam, "###,###,##0.00")
                              End If
                              
               'Provision de Riesgo Cambiario Crediticio ME - 14290105000000
               Case "ricmme":
                              If (g_rst_Princi!TIPMON = 2) Then
                                 r_arr_CtaRcd(r_int_Contad).CtaRcd_Import = Format(g_rst_Princi!PRVCAM * r_dbl_TipCam, "###,###,##0.00")
                              End If
               
                              
            End Select
            
            If r_arr_CtaRcd(r_int_Contad).CtaRcd_Import <> 0 Then
               
               'Insertando Registro de Saldos por Cliente
               g_str_Parame = ""
               g_str_Parame = g_str_Parame & "INSERT INTO TMP_SALRCD("
               g_str_Parame = g_str_Parame & "SALRCD_PERMES, "
               g_str_Parame = g_str_Parame & "SALRCD_PERANO, "
               g_str_Parame = g_str_Parame & "SALRCD_FERCRE, "
               g_str_Parame = g_str_Parame & "SALRCD_HORCRE, "
               g_str_Parame = g_str_Parame & "SALRCD_TERCRE, "
               g_str_Parame = g_str_Parame & "SALRCD_SALITE, "
               g_str_Parame = g_str_Parame & "SALRCD_TIPFOR, "
               g_str_Parame = g_str_Parame & "SALRCD_TIPINF, "
               g_str_Parame = g_str_Parame & "SALRCD_NUMSEC, "
               g_str_Parame = g_str_Parame & "SALRCD_TIPDOC, "
               g_str_Parame = g_str_Parame & "SALRCD_NUMDOC, "
               g_str_Parame = g_str_Parame & "SALRCD_CODOFI, "
               g_str_Parame = g_str_Parame & "SALRCD_UBIGEO, "
               g_str_Parame = g_str_Parame & "SALRCD_CTACTB, "
               g_str_Parame = g_str_Parame & "SALRCD_TIPCRE, "
               g_str_Parame = g_str_Parame & "SALRCD_MTOSDO, "
               g_str_Parame = g_str_Parame & "SALRCD_CONDIA, "
               g_str_Parame = g_str_Parame & "SALRCD_CONCTA, "
               g_str_Parame = g_str_Parame & "SALRCD_CONDIS) "
               
               
               g_str_Parame = g_str_Parame & "VALUES ("
               
               'Datos Basicos
               g_str_Parame = g_str_Parame & Mid(p_FecIni, 5, 2) & ", "
               g_str_Parame = g_str_Parame & Mid(p_FecIni, 1, 4) & ", "
               g_str_Parame = g_str_Parame & "'" & l_str_Fecha & "', "
               g_str_Parame = g_str_Parame & "'" & l_str_Hora & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
               
               'Nro de Iteracion
               g_str_Parame = g_str_Parame & r_int_SalIte & ", "
               
               'Tipo de Formulario 1 = Del Deudor / 2 = Totales de la Empresa
               g_str_Parame = g_str_Parame & 1 & ", "
               
               'Tipo de Informacion
               g_str_Parame = g_str_Parame & 2 & ", "
               
               'Nro de Secuencia
               g_str_Parame = g_str_Parame & "'" & Format(CStr(r_int_NumSec), "00000000") & "', "
               
               'Tipo de Documento
               g_str_Parame = g_str_Parame & g_rst_Princi!TDOCLI & ", "
               
               'Nro de Documento
               g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!NDOCLI) & "', "
               
               'Codigo de la Empresa Informante
               g_str_Parame = g_str_Parame & "'" & Format("0001", "0000") & "', "
               
               'Ubicacion Geografica de la Oficina de la empresa Informante (Dpto-Prov-Dist)
               g_str_Parame = g_str_Parame & "'" & Format("150131", "000000") & "', "
               
               'Codigo de Cuenta Contable
               g_str_Parame = g_str_Parame & "'" & Format(r_arr_CtaRcd(r_int_Contad).CtaRcd_NumCta, "00000000000000") & "', "
               
               'Tipo de Credito
               g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CLAPRD & "', "
            
               'Saldo
               g_str_Parame = g_str_Parame & r_arr_CtaRcd(r_int_Contad).CtaRcd_Import & ", "
               
               'Condicion en dias
               g_str_Parame = g_str_Parame & "'" & Format(Trim(g_rst_Princi!DIAMOR), "0000") & "', "
               
               'Condicion especial de la cuenta
               g_str_Parame = g_str_Parame & "'" & "02" & "', "
               
               'Condicion de disponibilidad
               g_str_Parame = g_str_Parame & "'" & "06" & "') "
                          
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                  Exit Sub
               End If
               
               r_int_SalIte = r_int_SalIte + 1
               
            End If
            
         Next r_int_Contad
                          
         r_int_NumSec = r_int_NumSec + 1
                  
         r_lng_NumReg = r_lng_NumReg + 1
                  
         g_rst_Princi.MoveNext
         DoEvents
                           
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
                  
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
   Else
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
 
      Screen.MousePointer = 0
 
      MsgBox "No se encontraron Datos registradas.", vbInformation, modgen_g_str_NomPlt
 
      Exit Sub
   End If
     

End Sub

Private Sub fs_Genera_ArcRcd(ByVal p_NomFil As String, ByVal p_FecIni As String, ByVal p_FecFin As String)

   Dim r_int_NumRes        As Integer
   Dim r_int_Contad        As Integer
   
   Dim r_str_NomRes        As String
   Dim r_str_CodSbs        As String
   Dim r_str_NumSec        As String
   Dim r_str_ExtSal        As String
   Dim r_str_ApePat        As String
   Dim r_str_ApeMat        As String
   Dim r_str_PriNom        As String
   Dim r_str_SegNom        As String
   Dim r_str_ApeCas        As String
   
   Dim r_rst_DesRcd        As ADODB.Recordset
   Dim r_rst_SalRcd        As ADODB.Recordset
   Dim r_rst_TotRcd        As ADODB.Recordset
   
   g_str_Parame = "SELECT * FROM MNT_EMPGRP WHERE "
   g_str_Parame = g_str_Parame & "EMPGRP_CODIGO = 000001 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_str_CodSbs = Trim(g_rst_Princi!EMPGRP_CODSBS)
   
   For r_int_Contad = Len(p_NomFil) To 1 Step -1
      If Mid(p_NomFil, r_int_Contad, 1) = "\" Then
         Exit For
      End If
   Next r_int_Contad

   r_str_NomRes = "C:\PruebaRCD\" & p_NomFil & Mid(p_FecIni, 1, 6) & "." & r_str_CodSbs
   
   'Creando Archivo de RCD
   r_int_NumRes = FreeFile
   
   Open r_str_NomRes For Output As r_int_NumRes
   
   Print #r_int_NumRes, "0106" & "01" & "00" & r_str_CodSbs & p_FecFin & "012" & "               " & "000000000000000"
   
   r_str_NumSec = 1
   
   g_str_Parame = "SELECT * FROM TMP_DESRCD "
   g_str_Parame = g_str_Parame & "ORDER BY DESRCD_APEPAT, DESRCD_APEMAT, DESRCD_APECAS, DESRCD_PRINOM, DESRCD_SEGNOM ASC "
      
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_DesRcd, 3) Then
      Exit Sub
   End If
      
   If Not (r_rst_DesRcd.BOF And r_rst_DesRcd.EOF) Then
      r_rst_DesRcd.MoveFirst
            
      Do While Not r_rst_DesRcd.EOF
      
               r_str_ApePat = ff_Caracter(r_rst_DesRcd!DESRCD_APEPAT)
               r_str_ApeMat = ff_Caracter(r_rst_DesRcd!DESRCD_APEMAT)
               r_str_PriNom = ff_Caracter(r_rst_DesRcd!DESRCD_PRINOM)
               r_str_SegNom = ff_Caracter(r_rst_DesRcd!DESRCD_SEGNOM)
               r_str_ApeCas = ff_Caracter(r_rst_DesRcd!DESRCD_APECAS)
               
                              
    '           Print #r_int_NumRes, r_rst_DesRcd!DESRCD_TIPFOR & r_rst_DesRcd!DESRCD_TIPINF & Format(r_str_NumSec, "00000000") & _
     '                               r_rst_DesRcd!DESRCD_CODSBS & r_rst_DesRcd!DESRCD_CODINT & r_rst_DesRcd!DESRCD_CODCIU & _
      '                              r_rst_DesRcd!DESRCD_CODOFI & r_rst_DesRcd!DESRCD_CODSOF & r_rst_DesRcd!DESRCD_TIPIDE & _
       '                             IIf(r_rst_DesRcd!DESRCD_TIPIDE = "T", r_rst_DesRcd!DESRCD_NUMTOM + r_rst_DesRcd!DESRCD_NUMFOL, r_rst_DesRcd!DESRCD_NUMPAR) & r_rst_DesRcd!DESRCD_TIPTRI & _
        '                            r_rst_DesRcd!DESRCD_DOCTRI & r_rst_DesRcd!DESRCD_TIPDOC & r_rst_DesRcd!DESRCD_NUMDOC & _
         '                           r_rst_DesRcd!DESRCD_TIPPER & r_rst_DesRcd!DESRCD_RESIDE & _
          '                          r_rst_DesRcd!DESRCD_CLADEU & r_rst_DesRcd!DESRCD_MAGSBS & r_rst_DesRcd!DESRCD_ACCINF & _
           '                         r_rst_DesRcd!DESRCD_RELLAB & r_rst_DesRcd!DESRCD_PAIRES & r_rst_DesRcd!DESRCD_TIPGEN & _
                                    r_rst_DesRcd!DESRCD_ESTCIV & r_rst_DesRcd!DESRCD_SIGLA & r_rst_DesRcd!DESRCD_APEPAT & _
            '                        r_rst_DesRcd!DESRCD_APEMAT & r_rst_DesRcd!DESRCD_APECAS & r_rst_DesRcd!DESRCD_PRINOM & _
             '                       r_rst_DesRcd!DESRCD_SEGNOM & r_rst_DesRcd!DESRCD_RIECAM & r_rst_DesRcd!DESRCD_INDATR & _
              '                      r_rst_DesRcd!DESRCD_CLAREP
                                               
               Print #r_int_NumRes, r_rst_DesRcd!DESRCD_TIPFOR & r_rst_DesRcd!DESRCD_TIPINF & Format(r_str_NumSec, "00000000") & _
                                    r_rst_DesRcd!DESRCD_CODSBS & r_rst_DesRcd!DESRCD_CODINT & Format(r_rst_DesRcd!DESRCD_CODCIU, "0000") & _
                                    r_rst_DesRcd!DESRCD_CODOFI & r_rst_DesRcd!DESRCD_CODSOF & r_rst_DesRcd!DESRCD_TIPTRI & Left(r_rst_DesRcd!DESRCD_DOCTRI, 10) + " " & r_rst_DesRcd!DESRCD_TIPIDE & _
                                    IIf(r_rst_DesRcd!DESRCD_TIPIDE = "T", r_rst_DesRcd!DESRCD_NUMTOM + r_rst_DesRcd!DESRCD_NUMFOL, r_rst_DesRcd!DESRCD_NUMPAR) & _
                                    r_rst_DesRcd!DESRCD_TIPDOC & r_rst_DesRcd!DESRCD_NUMDOC & _
                                    r_rst_DesRcd!DESRCD_TIPPER & r_rst_DesRcd!DESRCD_RESIDE & _
                                    r_rst_DesRcd!DESRCD_CLADEU & r_rst_DesRcd!DESRCD_MAGSBS & r_rst_DesRcd!DESRCD_ACCINF & _
                                    r_rst_DesRcd!DESRCD_RELLAB & r_rst_DesRcd!DESRCD_PAIRES & r_rst_DesRcd!DESRCD_TIPGEN & _
                                    r_rst_DesRcd!DESRCD_ESTCIV & r_rst_DesRcd!DESRCD_SIGLA & r_str_ApePat & _
                                    r_str_ApeMat & r_str_ApeCas & r_str_PriNom & _
                                    r_str_SegNom & r_rst_DesRcd!DESRCD_RIECAM & r_rst_DesRcd!DESRCD_INDATR & _
                                    r_rst_DesRcd!DESRCD_CLAREP
                                    
                                               
                                    
               g_str_Parame = "SELECT * FROM TMP_SALRCD WHERE "
               
               If (Trim(r_rst_DesRcd!DESRCD_TIPTRI) = 2 Or Trim(r_rst_DesRcd!DESRCD_TIPTRI) = 3) Then
                  g_str_Parame = g_str_Parame & "SALRCD_TIPDOC = '7' AND "
                  g_str_Parame = g_str_Parame & "SALRCD_NUMDOC = '" & Trim(r_rst_DesRcd!DESRCD_DOCTRI) & "' "
               Else
                  g_str_Parame = g_str_Parame & "SALRCD_TIPDOC = " & r_rst_DesRcd!DESRCD_TIPDOC & " AND "
                  g_str_Parame = g_str_Parame & "SALRCD_NUMDOC = '" & Trim(r_rst_DesRcd!DESRCD_NUMDOC) & "' "
               End If
               
               g_str_Parame = g_str_Parame & "ORDER BY  SALRCD_NUMSEC ASC "
                                    
               If Not gf_EjecutaSQL(g_str_Parame, r_rst_SalRcd, 3) Then
                  Exit Sub
               End If
                                                
               If Not (r_rst_SalRcd.BOF And r_rst_SalRcd.EOF) Then
                  
                  r_rst_SalRcd.MoveFirst
               
                  Do While Not r_rst_SalRcd.EOF
                                       
                     r_str_ExtSal = modtac_gs_Cadena_ExtSal(Format(CDbl(r_rst_SalRcd!SALRCD_MTOSDO), "###,###,##0.00"))
                                       
                     'r_str_ExtSal = GenNum(Format(r_rst_SalRcd!SALRCD_MTOSDO, "###,###,##0.00"))
                                       
                     'Print #r_int_NumRes, r_rst_SalRcd!SALRCD_TIPFOR & r_rst_SalRcd!SALRCD_TIPINF & Format(r_str_NumSec, "00000000") & _
                     '                     r_rst_SalRcd!SALRCD_CODOFI & r_rst_SalRcd!SALRCD_UBIGEO & r_rst_SalRcd!SALRCD_CTACTB & _
                     '                     r_rst_SalRcd!SALRCD_TIPCRE & Format(r_str_ExtSal, "000000000000000") & _
                     '                     r_rst_SalRcd!SALRCD_CONDIA & _
                     '                     r_rst_SalRcd!SALRCD_CONCTA & r_rst_SalRcd!SALRCD_CONDIS
                                          
                                          
                     Print #r_int_NumRes, r_rst_SalRcd!SALRCD_TIPFOR & r_rst_SalRcd!SALRCD_TIPINF & Format(r_str_NumSec, "00000000") & _
                                          r_rst_SalRcd!SALRCD_CODOFI & r_rst_SalRcd!SALRCD_UBIGEO & r_rst_SalRcd!SALRCD_CTACTB & _
                                          r_rst_SalRcd!SALRCD_TIPCRE & Format(r_str_ExtSal, "000000000000000") & _
                                          r_rst_SalRcd!SALRCD_CONDIA & _
                                          r_rst_SalRcd!SALRCD_CONDIS & r_rst_SalRcd!SALRCD_CONCTA
                                          
                                         
                  r_rst_SalRcd.MoveNext
                  DoEvents
                  
                  Loop
               End If
               
               r_str_NumSec = r_str_NumSec + 1
               
            r_rst_DesRcd.MoveNext
            DoEvents
         
      Loop
   End If
   
   Print #r_int_NumRes, "21" & Format(r_str_NumSec, "00000000")
   r_str_NumSec = r_str_NumSec + 1
   
   g_str_Parame = "SELECT DISTINCT(SALRCD_CTACTB) AS CUENTA, SUM(SALRCD_MTOSDO) AS SALDO, MAX(SALRCD_TIPCRE) AS TIPCRE, MAX(SALRCD_CONDIA) AS CONDIA FROM TMP_SALRCD "
   g_str_Parame = g_str_Parame & "GROUP BY SALRCD_CTACTB, SALRCD_TIPCRE"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_TotRcd, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_TotRcd.BOF And r_rst_TotRcd.EOF) Then
                  
      r_rst_TotRcd.MoveFirst
      
      Do While Not r_rst_TotRcd.EOF
      
         r_str_ExtSal = modtac_gs_Cadena_ExtSal(Format(CDbl(r_rst_TotRcd!SALDO), "###,###,##0.00"))
         'r_str_ExtSal = GenNum(Format(r_rst_TotRcd!SALDO, "###,###,##0.00"))
                           
         Print #r_int_NumRes, 2 & 2 & Format(r_str_NumSec, "00000000") & "0000000000" & _
                              r_rst_TotRcd!Cuenta & r_rst_TotRcd!TIPCRE & Format(r_str_ExtSal, "000000000000000") & _
                              Format(r_rst_TotRcd!CONDIA, "0000")
                              
         r_rst_TotRcd.MoveNext
         DoEvents
                 
      Loop
   End If
   
   r_rst_TotRcd.Close
   Set r_rst_TotRcd = Nothing
   
   'r_rst_SalRcd.Close
   'Set r_rst_SalRcd = Nothing
   
   r_rst_DesRcd.Close
   Set r_rst_DesRcd = Nothing
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   'Cerrando Archivo de RCD
   Close #r_int_NumRes


End Sub


Public Function ff_ComCie(ByVal p_FecIni As String, ByVal p_ndocli As String, Optional ByRef p_Clacli As String, Optional ByRef p_ClaPrv As String, Optional ByRef p_ClaPrd As String, Optional ByRef p_DiaMor As String) As String
      
   g_str_Parame = "SELECT DISTINCT(COMCIE_NDOCLI), MAX(COMCIE_TDOCLI), MAX(COMCIE_CLACLI) AS CLACLI, MAX(COMCIE_CLAPRV) AS CLAPRV, MAX(COMCIE_CLAPRD) AS CLAPRD,MAX(COMCIE_DIAMOR) AS DIAMOR FROM CRE_COMCIE WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & Mid(p_FecIni, 1, 4) & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & Mid(p_FecIni, 5, 2) & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_TDOCLI = 7 AND "
   g_str_Parame = g_str_Parame & "COMCIE_NDOCLI = '" & p_ndocli & "' "
   g_str_Parame = g_str_Parame & "GROUP BY COMCIE_NDOCLI "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      
      p_Clacli = g_rst_Listas!CLACLI
      p_ClaPrv = g_rst_Listas!CLAPRV
      p_ClaPrd = g_rst_Listas!CLAPRD
      p_DiaMor = g_rst_Listas!DIAMOR
   
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
   
End Function

'Public Function ff_ExtSal(ByVal p_Saldos As Double) As String
'
'   Dim r_int_Count As Integer
'
'   p_Saldos = Trim(CStr(p_Saldos))
'
'   r_int_Count = 1
'
'   Do While Len(Trim(Mid(p_Saldos, r_int_Count, 1))) > 0
'
'      If Trim(Mid(p_Saldos, r_int_Count, 1)) <> "." Then
'         ff_ExtSal = ff_ExtSal + Mid(p_Saldos, r_int_Count, 1)
'         r_int_Count = r_int_Count + 1
'      Else
'         If Mid(Right(Trim(p_Saldos), 2), 1, 1) <> "." Then
'            ff_ExtSal = ff_ExtSal + Right(Trim(p_Saldos), 2)
'            Exit Function
'         Else
'            ff_ExtSal = ff_ExtSal + Right(Trim(p_Saldos), 1) + "0"
'            Exit Function
'         End If
'      End If
'   Loop
'
'   ff_ExtSal = ff_ExtSal + "00"
   
'End Function

Public Function ff_Nombre(ByVal p_Nombre As String, Optional ByRef p_PriNom As String, Optional ByRef p_SegNom As String) As String
   
   Dim r_int_Count As Integer
      
   p_Nombre = Trim(p_Nombre)
   p_PriNom = " "
   p_SegNom = " "
         
   r_int_Count = 1
   
   Do While Len(Mid(p_Nombre, r_int_Count, 1)) > 0
         
      If Mid(p_Nombre, r_int_Count, 1) <> " " Then
         p_PriNom = p_PriNom + Mid(p_Nombre, r_int_Count, 1)
         r_int_Count = r_int_Count + 1
      Else
         p_SegNom = Mid(p_Nombre, r_int_Count + 1, Len(p_Nombre))
         Exit Do
      End If
   
   Loop
   
End Function

Public Function ff_Caracter(ByVal p_Nombre As String) As String
   
   Dim r_int_Count As Integer
      
   p_Nombre = p_Nombre
   
   r_int_Count = 1
   
   Do While Len(Mid(p_Nombre, r_int_Count, 1)) > 0
         
      If Mid(p_Nombre, r_int_Count, 1) <> "Ñ" Then
         ff_Caracter = ff_Caracter + Mid(p_Nombre, r_int_Count, 1)
      Else
         ff_Caracter = ff_Caracter + "#"
      End If
      
      'If Mid(p_Nombre, r_int_Count + 1, 1) = " " Then
      '   Exit Function
      'End If
      
      r_int_Count = r_int_Count + 1
   
   Loop
   
End Function

Public Function ff_ConHip() As Integer
   
   ff_ConHip = 0
   
   g_str_Parame = "SELECT NVL(COUNT(*),0) AS TOTREG FROM CRE_HIPCIE WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00") & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
            
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      
      ff_ConHip = g_rst_Listas!TOTREG
   
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
      
End Function

Public Function ff_ConCom() As Integer
   
   ff_ConCom = 0
   
   g_str_Parame = "SELECT NVL(COUNT(DISTINCT(COMCIE_NDOCLI)),0) AS TOTREG FROM CRE_COMCIE WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00") & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
            
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      
      ff_ConCom = g_rst_Listas!TOTREG
   
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
      
End Function

Public Function ff_LinGar(ByVal p_NumOpe As String) As String
   
   g_str_Parame = "SELECT HIPMAE_GARLIN FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = " & p_NumOpe & " "
            
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      
      ff_LinGar = g_rst_Listas!HIPMAE_GARLIN
   
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
      
End Function

Private Function GenNum(ByVal p_Numero As String) As String
   Dim l_cadena As String
   Dim l_contad As Integer
   Dim l_longit As Integer
   
   l_longit = Len(p_Numero)
      
   For l_contad = 1 To l_longit Step 1
      If Mid(p_Numero, l_contad, 1) <> "." Then
         GenNum = GenNum + Mid(p_Numero, l_contad, 1)
      End If
   Next l_contad

   'GenNum = Format(l_cadena, "000000000000000")

End Function




