VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_CieCre_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3840
   ClientLeft      =   7710
   ClientTop       =   5085
   ClientWidth     =   7935
   Icon            =   "GesCtb_frm_162.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3855
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7965
      _Version        =   65536
      _ExtentX        =   14049
      _ExtentY        =   6800
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
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
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
            Height          =   315
            Left            =   660
            TabIndex        =   17
            Top             =   30
            Width           =   3045
            _Version        =   65536
            _ExtentX        =   5371
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cierre de Cartera de Creditos"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   375
            Left            =   660
            TabIndex        =   18
            Top             =   240
            Width           =   2355
            _Version        =   65536
            _ExtentX        =   4154
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Procesos"
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
            Picture         =   "GesCtb_frm_162.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   6
         Top             =   780
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
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
         Begin VB.CommandButton cmd_AnuPro 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_162.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7230
            Picture         =   "GesCtb_frm_162.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_162.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Procesar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   7
         Top             =   1470
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   1349
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
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   6285
         End
         Begin Threed.SSPanel pnl_Period 
            Height          =   315
            Left            =   1530
            TabIndex        =   9
            Top             =   390
            Width           =   6285
            _Version        =   65536
            _ExtentX        =   11086
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label lbl_NomEti 
            Caption         =   "Período:"
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   10
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   735
         Left            =   30
         TabIndex        =   11
         Top             =   3060
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   1296
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
            TabIndex        =   12
            Top             =   360
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SSPanel2"
            ForeColor       =   0
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
         Begin VB.Label lbl_NomPro 
            Caption         =   "Cierre de Créditos Hipotecarios"
            Height          =   255
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   5500
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   735
         Left            =   30
         TabIndex        =   14
         Top             =   2280
         Width           =   7875
         _Version        =   65536
         _ExtentX        =   13891
         _ExtentY        =   1296
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
         Begin Threed.SSPanel pnl_BarTot 
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   360
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "SSPanel2"
            ForeColor       =   0
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
         Begin VB.Label lbl_Erique 
            Caption         =   "Proceso Total:"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_CieCre_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Empres()      As moddat_tpo_Genera

Private Sub cmb_Empres_Click()
   If cmb_Empres.ListIndex > -1 Then
      Screen.MousePointer = 11
      
      pnl_Period.Caption = moddat_gf_ConsultaPerMesActivo(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, 2, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
      
      'Verificar procesos
      g_str_Parame = "SELECT NVL(COUNT(*),0) AS TOTREG FROM CRE_HIPCIE WHERE "
      g_str_Parame = g_str_Parame & "HIPCIE_CODEMP = '" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "' AND "
      g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & CStr(modctb_int_PerMes) & " AND "
      g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & CStr(modctb_int_PerAno) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi!TOTREG = 0 Then
         cmd_Proces.Enabled = True
         cmd_AnuPro.Enabled = False
         Call gs_SetFocus(cmd_Proces)
      Else
         cmd_Proces.Enabled = False
         cmd_AnuPro.Enabled = True
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmd_AnuPro_Click()
   Dim r_lng_TotErr     As Long

   If cmb_Empres.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Empres)
      Exit Sub
   End If
   
   If MsgBox("Está seguro de ejecutar el proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   cmd_AnuPro.Enabled = False
   Screen.MousePointer = 11
   Call modprc_ctbp1002("CTBP1002", Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_TotErr, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, pnl_BarPro)
   Screen.MousePointer = 0
   cmd_Proces.Enabled = True
End Sub

Private Sub cmd_Proces_Click()
   Dim r_lng_TotErr     As Long

   If cmb_Empres.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Empres)
      Exit Sub
   End If
   
   If MsgBox("Está seguro de ejecutar el proceso de Cierre de Cartera?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   cmd_Proces.Enabled = False
   Screen.MousePointer = 11
   pnl_BarTot.FloodPercent = 0
   
   'Cierre de Créditos Hipotecarios
   lbl_NomPro.Caption = "Cierre de Créditos Hipotecarios:"
   Call modprc_ctbp1011(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, modctb_int_PerMes, modctb_int_PerAno, modctb_str_FecFin, modctb_str_FecIni, pnl_BarPro)
   pnl_BarTot.FloodPercent = 20
   DoEvents: DoEvents
   
'''   'Cierre de Crèditos Comerciales
'''   lbl_NomPro.Caption = "Cierre de Créditos Comerciales:"
'''   Call gs_Datos_CreCom(modctb_int_PerMes, modctb_int_PerAno, modctb_str_FecFin)
'''   pnl_BarTot.FloodPercent = 30
'''   DoEvents: DoEvents
   
   'Clasificación de Clientes
   lbl_NomPro.Caption = "Clasificación de Clientes:"
   Call modprc_ctbp1008(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, modctb_int_PerMes, modctb_int_PerAno, modctb_str_FecFin, pnl_BarPro)
   pnl_BarTot.FloodPercent = 30 '40
   DoEvents: DoEvents
   
   'Cálculo de Provisiones
   lbl_NomPro.Caption = "Cálculo de Provisiones:"
   Call modprc_ctbp1009(l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, modctb_int_PerMes, modctb_int_PerAno, modctb_str_FecFin, pnl_BarPro)
   pnl_BarTot.FloodPercent = 40 '50
   DoEvents: DoEvents
   
   'Llenado Credito_Cierre_Finmes (Temporal)
   Call gs_Cargar_CreCie(modctb_int_PerMes, modctb_int_PerAno)
'   Call gs_Cargar_CreCom(modctb_int_PerMes, modctb_int_PerAno)
   pnl_BarTot.FloodPercent = 50 '60
   DoEvents: DoEvents
   
   'Mantenedor de Cuentas RCD
   Call gs_Datos_CueRcd(modctb_int_PerMes, modctb_int_PerAno)
   pnl_BarTot.FloodPercent = 60 '70
   DoEvents: DoEvents
   
   'Mantenedor de Personal
   Call gs_Datos_Person(modctb_int_PerMes, modctb_int_PerAno)
   pnl_BarTot.FloodPercent = 70 '80
   DoEvents: DoEvents
   
   'Cierre de Cartas Fianzas (tpr_cafcie)
   lbl_NomPro.Caption = "Cierre de Cartas Fianza:"
   Call gs_Datos_CafCie(modctb_int_PerMes, modctb_int_PerAno)
   pnl_BarTot.FloodPercent = 80 '90
   DoEvents: DoEvents
   
    'Cierre de Crèditos Comerciales
   lbl_NomPro.Caption = "Cierre de Créditos Comerciales:"
   Call gs_Datos_CreCom(modctb_int_PerMes, modctb_int_PerAno, modctb_str_FecFin)
   pnl_BarTot.FloodPercent = 90 '30
   DoEvents: DoEvents
     
   'Llenado Credito_Cierre_Finmes (Temporal)
   Call gs_Cargar_CreCom(modctb_int_PerMes, modctb_int_PerAno)
   pnl_BarTot.FloodPercent = 92
   DoEvents: DoEvents
   
   
   'Cierre de Cartas Fianzas - Venta y Patrimonio (CTB_VTAPAT)
   lbl_NomPro.Caption = "Cierre de Cartas Fianza Venta y Patrimonio:"
   Call gs_CafCie_VtaPat(modctb_int_PerMes, modctb_int_PerAno)
   pnl_BarTot.FloodPercent = 95
   DoEvents: DoEvents
   
   'Cierre de Bienes adjudicados (CRE_BIEADJ)
   lbl_NomPro.Caption = "Cierre de Bienes Adjudicados:"
   Call gs_Datos_BieAdj(modctb_int_PerMes, modctb_int_PerAno)
   pnl_BarTot.FloodPercent = 100
   DoEvents: DoEvents

   Screen.MousePointer = 0
   MsgBox "Proceso concluído.", vbInformation, modgen_g_str_NomPlt
   cmd_AnuPro.Enabled = True
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
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
End Sub

Private Sub fs_Limpia()
   cmd_Proces.Enabled = False
   cmd_AnuPro.Enabled = False
   cmb_Empres.ListIndex = 0
   pnl_BarPro.FloodPercent = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Public Sub gs_Datos_CueRcd(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
   g_str_Parame = ""
   g_str_Parame = "SELECT * FROM CTB_CUERCD WHERE "
   g_str_Parame = g_str_Parame & "CUERCD_PERMES = " & IIf(p_PerMes - 1 = 0, 12, p_PerMes - 1) & " AND "
   g_str_Parame = g_str_Parame & "CUERCD_PERANO = " & IIf(p_PerMes - 1 = 0, p_PerAno - 1, p_PerAno) & " "
   g_str_Parame = g_str_Parame & "ORDER BY CUERCD_CTACTB ASC "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CTB_CUERCD("
      g_str_Parame = g_str_Parame & "CUERCD_PERMES, "
      g_str_Parame = g_str_Parame & "CUERCD_PERANO, "
      g_str_Parame = g_str_Parame & "CUERCD_CTACTB, "
      g_str_Parame = g_str_Parame & "CUERCD_DESVAR, "
      g_str_Parame = g_str_Parame & "CUERCD_DESCRI) "
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & p_PerMes & ", "
      g_str_Parame = g_str_Parame & p_PerAno & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!CUERCD_CTACTB) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!CUERCD_DESVAR) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!CUERCD_DESCRI) & "') "
               
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub gs_Datos_Person(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
   g_str_Parame = ""
   g_str_Parame = "SELECT * FROM CTB_USUSBS WHERE "
   g_str_Parame = g_str_Parame & "USUSBS_PERMES = " & IIf(p_PerMes - 1 = 0, 12, p_PerMes - 1) & " AND "
   g_str_Parame = g_str_Parame & "USUSBS_PERANO = " & IIf(p_PerMes - 1 = 0, p_PerAno - 1, p_PerAno) & " "
   g_str_Parame = g_str_Parame & "ORDER BY USUSBS_CODUSU ASC "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CTB_USUSBS("
      g_str_Parame = g_str_Parame & "USUSBS_PERMES, "
      g_str_Parame = g_str_Parame & "USUSBS_PERANO, "
      g_str_Parame = g_str_Parame & "USUSBS_CODUSU, "
      g_str_Parame = g_str_Parame & "USUSBS_APEPAT, "
      g_str_Parame = g_str_Parame & "USUSBS_APEMAT, "
      g_str_Parame = g_str_Parame & "USUSBS_NOMBRE, "
      g_str_Parame = g_str_Parame & "USUSBS_TIPPER ) "
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & p_PerMes & ", "
      g_str_Parame = g_str_Parame & p_PerAno & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!USUSBS_CODUSU) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!USUSBS_APEPAT) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!USUSBS_APEMAT) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!USUSBS_NOMBRE) & "', "
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!USUSBS_TIPPER) & "') "
               
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub gs_Datos_CreCom(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecFin As String)
Dim r_arr_TipPrv()         As modprc_g_tpo_TipPrv
Dim r_arr_LogPro()         As modprc_g_tpo_LogPro
Dim r_arr_DetGar()         As modprc_g_tpo_DetGar
Dim r_str_FecPro           As String
Dim r_dbl_TipCam_Dol       As Double
Dim r_dbl_TipCam           As Double
Dim r_dbl_PrvGen           As Double
Dim r_dbl_PrvEsp           As Double
Dim r_dbl_PrvCic           As Double
Dim r_int_ClaGar           As Integer
Dim r_dbl_PrvCam           As Double
Dim r_int_NueCre           As Integer
   
   r_dbl_PrvGen = 0
   r_dbl_PrvEsp = 0
   r_dbl_PrvCic = 0
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_FecFin), "yyyymmdd"), 2)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * FROM TPR_CAFCIE "
   g_str_Parame = g_str_Parame & "  WHERE CAFCIE_PERMES = " & p_PerMes & " "
   g_str_Parame = g_str_Parame & "    AND CAFCIE_PERANO = " & p_PerAno & " "
   g_str_Parame = g_str_Parame & "    AND CAFCIE_CODPRD = '008' "
   g_str_Parame = g_str_Parame & "  ORDER BY CAFCIE_NUMREF ASC "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
      
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         
         If g_rst_Listas!CAFCIE_TIPEMP = 4 Then
            r_int_NueCre = 7
         ElseIf g_rst_Listas!CAFCIE_TIPEMP = 3 Then
            r_int_NueCre = 8
         ElseIf g_rst_Listas!CAFCIE_TIPEMP = 2 Then
            r_int_NueCre = 9
         End If
   
         'Leer Tablas de Provisiones para Créditos Hipotecarios
         modprc_g_str_CadEje = "SELECT * FROM CTB_TIPPRV WHERE TIPPRV_CLACRE = '" & r_int_NueCre & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CTB_TIPPRV.")
            modprc_g_rst_Princi.Close
            Set modprc_g_rst_Princi = Nothing
            Exit Sub
         End If
         
         modprc_g_rst_Princi.MoveFirst
         ReDim r_arr_TipPrv(0)
         
         Do While Not modprc_g_rst_Princi.EOF
            ReDim Preserve r_arr_TipPrv(UBound(r_arr_TipPrv) + 1)
            r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_TipPrv = CInt(modprc_g_rst_Princi!TipPrv_TipPrv)
            r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_CodCla = CInt(modprc_g_rst_Princi!TIPPRV_CLFCRE)
            r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_ClaGar = CInt(modprc_g_rst_Princi!TipPrv_ClaGar)
            r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_Porcen = modprc_g_rst_Princi!TipPrv_Porcen
            modprc_g_rst_Princi.MoveNext
         Loop
         
         modprc_g_rst_Princi.Close
         Set modprc_g_rst_Princi = Nothing
         
         'Leer Tabla de Garantías CTB_DETGAR
         modprc_g_str_CadEje = "SELECT * FROM CTB_DETGAR "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CTB_DETGAR.")
            modprc_g_rst_Princi.Close
            Set modprc_g_rst_Princi = Nothing
            Exit Sub
         End If
         
         modprc_g_rst_Princi.MoveFirst
         ReDim r_arr_DetGar(0)
         
         Do While Not modprc_g_rst_Princi.EOF
            ReDim Preserve r_arr_DetGar(UBound(r_arr_DetGar) + 1)
            r_arr_DetGar(UBound(r_arr_DetGar)).DetGar_Codigo = CInt(modprc_g_rst_Princi!DetGar_Codigo)
            r_arr_DetGar(UBound(r_arr_DetGar)).DetGar_ClaGar = CInt(modprc_g_rst_Princi!DetGar_ClaGar)
            modprc_g_rst_Princi.MoveNext
         Loop
         
         modprc_g_rst_Princi.Close
         Set modprc_g_rst_Princi = Nothing
         r_int_ClaGar = 1
         
         If g_rst_Listas!CAFCIE_FIAMON = 1 Or g_rst_Listas!CAFCIE_FIAMON = 2 Then
            r_dbl_TipCam = r_dbl_TipCam_Dol
         Else
            r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, g_rst_Listas!CAFCIE_FIAMON, Format(CDate(p_FecFin), "yyyymmdd"), 2)
         End If
            
         If g_rst_Listas!CAFCIE_CLAPRV = 0 Then
            'Calculando Provisión Generica
            r_dbl_PrvGen = modprc_gf_PorcenProv(r_arr_TipPrv, 1, g_rst_Listas!CAFCIE_CLAPRV, r_int_ClaGar) * (g_rst_Listas!CAFCIE_SALCAP) / 100
                                                      
            'Calculando Provisión Pro-Ciclica
            'r_dbl_PrvCic = modprc_gf_PorcenProv(r_arr_TipPrv, 3, g_rst_Listas!CAFCIE_CLAPRV, r_int_ClaGar) * (g_rst_Listas!CAFCIE_SALCAP) / 100
            r_dbl_PrvCic = 0
            
            If g_rst_Listas!CAFCIE_FIAMON = 2 Then
               'Calculando Provisión Riesgo Cambiario
               r_dbl_PrvCam = modprc_gf_PorcenProv(r_arr_TipPrv, 4, g_rst_Listas!CAFCIE_CLAPRV, r_int_ClaGar) * (g_rst_Listas!CAFCIE_SALCAP) / 100
            End If
         Else
            'Calculando Provisión Específica
            r_dbl_PrvEsp = modprc_gf_PorcenProv(r_arr_TipPrv, 2, g_rst_Listas!CAFCIE_CLAPRV, r_int_ClaGar) * (g_rst_Listas!CAFCIE_SALCAP) / 100
         End If
            
         r_dbl_PrvGen = CDbl(Format(r_dbl_PrvGen, "######0.00"))
         r_dbl_PrvEsp = CDbl(Format(r_dbl_PrvEsp, "######0.00"))
         r_dbl_PrvCic = CDbl(Format(r_dbl_PrvCic, "######0.00"))
         r_dbl_PrvCam = CDbl(Format(r_dbl_PrvCam, "######0.00"))
         
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE TPR_CAFCIE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET CAFCIE_FECCIE = " & Format(CDate(r_str_FecPro), "yyyymmdd") & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       CAFCIE_TIPCAM = " & CStr(r_dbl_TipCam) & ", "
         If g_rst_Listas!CAFCIE_DIAMOR = 0 Then
            modprc_g_str_CadEje = modprc_g_str_CadEje & "    CAFCIE_PRVGEN = " & CStr(r_dbl_PrvGen) & ", "
         Else
            modprc_g_str_CadEje = modprc_g_str_CadEje & "    CAFCIE_PRVESP = " & CStr(r_dbl_PrvEsp) & ", "
         End If
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       CAFCIE_PRVCIC = " & CStr(r_dbl_PrvCic) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       CAFCIE_NUECRE = " & CStr(r_int_NueCre) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       CAFCIE_PRVCAM = " & CStr(r_dbl_PrvCam) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE CAFCIE_PERMES = " & CStr(p_PerMes) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       CAFCIE_PERANO = " & CStr(p_PerAno) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       CAFCIE_NUMREF = " & g_rst_Listas!CAFCIE_NUMREF & " "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, g_rst_Genera, 2) Then
            Exit Sub
         End If
            
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub
Public Sub gs_Datos_CreCom_old(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_FecFin As String)
Dim r_arr_TipPrv()         As modprc_g_tpo_TipPrv
Dim r_arr_LogPro()         As modprc_g_tpo_LogPro
Dim r_arr_DetGar()         As modprc_g_tpo_DetGar
Dim r_str_FecPro           As String
Dim r_dbl_TipCam_Dol       As Double
Dim r_dbl_TipCam           As Double
Dim r_dbl_PrvGen           As Double
Dim r_dbl_PrvEsp           As Double
Dim r_dbl_PrvCic           As Double
Dim r_int_ClaGar           As Integer
Dim r_dbl_PrvCam           As Double
   
   r_dbl_PrvGen = 0
   r_dbl_PrvEsp = 0
   r_dbl_PrvCic = 0
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Obteniendo Tipo de Cambio de Cierre
   r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_FecFin), "yyyymmdd"), 2)
               
   g_str_Parame = "SELECT * FROM CRE_COMCIE WHERE COMCIE_PERMES = " & p_PerMes & " AND COMCIE_PERANO = " & p_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY COMCIE_NUMOPE ASC "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
      
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
      
         'Leer Tablas de Provisiones para Créditos Hipotecarios
         modprc_g_str_CadEje = "SELECT * FROM CTB_TIPPRV WHERE TIPPRV_CLACRE = '" & g_rst_Listas!COMCIE_NUECRE & "' "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CTB_TIPPRV.")
            modprc_g_rst_Princi.Close
            Set modprc_g_rst_Princi = Nothing
            Exit Sub
         End If
         
         modprc_g_rst_Princi.MoveFirst
         ReDim r_arr_TipPrv(0)
         
         Do While Not modprc_g_rst_Princi.EOF
            ReDim Preserve r_arr_TipPrv(UBound(r_arr_TipPrv) + 1)
            r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_TipPrv = CInt(modprc_g_rst_Princi!TipPrv_TipPrv)
            r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_CodCla = CInt(modprc_g_rst_Princi!TIPPRV_CLFCRE)
            r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_ClaGar = CInt(modprc_g_rst_Princi!TipPrv_ClaGar)
            r_arr_TipPrv(UBound(r_arr_TipPrv)).TipPrv_Porcen = modprc_g_rst_Princi!TipPrv_Porcen
            modprc_g_rst_Princi.MoveNext
         Loop
         
         modprc_g_rst_Princi.Close
         Set modprc_g_rst_Princi = Nothing
         
         'Leer Tabla de Garantías CTB_DETGAR
         modprc_g_str_CadEje = "SELECT * FROM CTB_DETGAR "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
            r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, "Error al Leer Tabla CTB_DETGAR.")
            modprc_g_rst_Princi.Close
            Set modprc_g_rst_Princi = Nothing
            Exit Sub
         End If
         
         modprc_g_rst_Princi.MoveFirst
         ReDim r_arr_DetGar(0)
         
         Do While Not modprc_g_rst_Princi.EOF
            ReDim Preserve r_arr_DetGar(UBound(r_arr_DetGar) + 1)
            r_arr_DetGar(UBound(r_arr_DetGar)).DetGar_Codigo = CInt(modprc_g_rst_Princi!DetGar_Codigo)
            r_arr_DetGar(UBound(r_arr_DetGar)).DetGar_ClaGar = CInt(modprc_g_rst_Princi!DetGar_ClaGar)
            modprc_g_rst_Princi.MoveNext
         Loop
         
         modprc_g_rst_Princi.Close
         Set modprc_g_rst_Princi = Nothing
         r_int_ClaGar = 1
         
         If g_rst_Listas!COMCIE_TIPMON = 1 Or g_rst_Listas!COMCIE_TIPMON = 2 Then
            r_dbl_TipCam = r_dbl_TipCam_Dol
         Else
            r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, g_rst_Listas!COMCIE_TIPMON, Format(CDate(p_FecFin), "yyyymmdd"), 2)
         End If
            
         If g_rst_Listas!COMCIE_CLAPRV = 0 Then
            'Calculando Provisión Generica
            r_dbl_PrvGen = modprc_gf_PorcenProv(r_arr_TipPrv, 1, g_rst_Listas!COMCIE_CLAPRV, r_int_ClaGar) * (g_rst_Listas!COMCIE_SALCAP) / 100
                                          
            'Calculando Provisión Pro-Ciclica
            r_dbl_PrvCic = modprc_gf_PorcenProv(r_arr_TipPrv, 3, g_rst_Listas!COMCIE_CLAPRV, r_int_ClaGar) * (g_rst_Listas!COMCIE_SALCAP) / 100
            
            If g_rst_Listas!COMCIE_TIPMON = 2 Then
               'Calculando Provisión Riesgo Cambiario
               r_dbl_PrvCam = modprc_gf_PorcenProv(r_arr_TipPrv, 4, g_rst_Listas!COMCIE_CLAPRV, r_int_ClaGar) * (g_rst_Listas!COMCIE_SALCAP) / 100
            End If
         Else
            'Calculando Provisión Específica
            r_dbl_PrvEsp = modprc_gf_PorcenProv(r_arr_TipPrv, 2, g_rst_Listas!COMCIE_CLAPRV, r_int_ClaGar) * (g_rst_Listas!COMCIE_SALCAP) / 100
         End If
            
         r_dbl_PrvGen = CDbl(Format(r_dbl_PrvGen, "######0.00"))
         r_dbl_PrvEsp = CDbl(Format(r_dbl_PrvEsp, "######0.00"))
         r_dbl_PrvCic = CDbl(Format(r_dbl_PrvCic, "######0.00"))
         r_dbl_PrvCam = CDbl(Format(r_dbl_PrvCam, "######0.00"))
         
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_COMCIE "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET COMCIE_FECCIE = " & Format(CDate(r_str_FecPro), "yyyymmdd") & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       COMCIE_TIPCAM = " & CStr(r_dbl_TipCam) & ", "
         If g_rst_Listas!COMCIE_DIAMOR = 0 Then
            modprc_g_str_CadEje = modprc_g_str_CadEje & "    COMCIE_PRVGEN = " & CStr(r_dbl_PrvGen) & ", "
         Else
            modprc_g_str_CadEje = modprc_g_str_CadEje & "    COMCIE_PRVESP = " & CStr(r_dbl_PrvEsp) & ", "
         End If
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       COMCIE_PRVCIC = " & CStr(r_dbl_PrvCic) & ", "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       COMCIE_PRVCAM = " & CStr(r_dbl_PrvCam) & " "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE COMCIE_PERMES = " & CStr(p_PerMes) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       COMCIE_PERANO = " & CStr(p_PerAno) & " AND "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "       COMCIE_NUMOPE = " & g_rst_Listas!COMCIE_NUMOPE & " "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, g_rst_Genera, 2) Then
            Exit Sub
         End If
            
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub gs_Cargar_CreCie(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
         
   g_str_Parame = "SELECT * FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & p_PerMes & " AND HIPCIE_PERANO = " & p_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_NUMOPE ASC "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   g_rst_Listas.MoveFirst
   Do While Not g_rst_Listas.EOF
   
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CREDITO_CIERRE_FINMES("
      g_str_Parame = g_str_Parame & "   CREDITO, "
      g_str_Parame = g_str_Parame & "   MES, "
      g_str_Parame = g_str_Parame & "   ANO, "
      g_str_Parame = g_str_Parame & "   COD_MONEDA, "
      g_str_Parame = g_str_Parame & "   CLIENTE, "
      g_str_Parame = g_str_Parame & "   PRODUCTO, "
      g_str_Parame = g_str_Parame & "   FLAG_ESTADO_CRED, "
      g_str_Parame = g_str_Parame & "   FLAG_ESTADO_ANT, "
      g_str_Parame = g_str_Parame & "   TASA, "
      g_str_Parame = g_str_Parame & "   CAPITAL_APROBADO, "
      g_str_Parame = g_str_Parame & "   CAPITAL_DESEMBOLSADO, "
      g_str_Parame = g_str_Parame & "   CAPITAL_AMORTIZADO, "
      g_str_Parame = g_str_Parame & "   INTERES, "
      g_str_Parame = g_str_Parame & "   INTERES_COMP, "
      g_str_Parame = g_str_Parame & "   INTERES_MOR, "
      g_str_Parame = g_str_Parame & "   OTROS_CARGOS, "
      g_str_Parame = g_str_Parame & "   SALDO_GARANTIA_MN, "
      g_str_Parame = g_str_Parame & "   SALDO_GARANTIA_ME, "
      g_str_Parame = g_str_Parame & "   DIAS_MOROSIDAD, "
      g_str_Parame = g_str_Parame & "   CUOTAS_ATRAZADAS, "
      g_str_Parame = g_str_Parame & "   CAPITAL_VENCIDO, "
      g_str_Parame = g_str_Parame & "   FECHA_ULT_MOV, "
      g_str_Parame = g_str_Parame & "   FECHA_VENCIMIENTO, "
      g_str_Parame = g_str_Parame & "   ACT_ECONOMICA, "
      g_str_Parame = g_str_Parame & "   SALDO_NO_CONCESIONAL, "
      g_str_Parame = g_str_Parame & "   FECHA_APROBACION, "
      g_str_Parame = g_str_Parame & "   CAPITAL_INTERES, "
      g_str_Parame = g_str_Parame & "   CAPITAL_AMORTIZADO_MES, "
      g_str_Parame = g_str_Parame & "   CIIU, "
      g_str_Parame = g_str_Parame & "   FECHA_DESEMBOLSO, "
      g_str_Parame = g_str_Parame & "   SALDO_CONCESIONAL, "
      g_str_Parame = g_str_Parame & "   TASA_COSTO_EFECTIVO) "
      
      g_str_Parame = g_str_Parame & "VALUES ("
      
      'Operacion
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!HIPCIE_NUMOPE) & "', "
      
      'Mes
      g_str_Parame = g_str_Parame & p_PerMes & ", "
      
      'Año
      g_str_Parame = g_str_Parame & p_PerAno & ", "
      
      'Moneda
      g_str_Parame = g_str_Parame & "'" & Format(Trim(g_rst_Listas!HIPCIE_TIPMON), "000") & "', "
      
      'Documento identificacion del cliente
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!HIPCIE_TDOCLI) & Trim(g_rst_Listas!HIPCIE_NDOCLI) & "', "
      
      'Producto
      g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!HIPCIE_CODPRD) & "', "
      
      'Clasificacion del cliente
      If Trim(g_rst_Listas!HIPCIE_SITCRE) = 4 Then
         If Trim(g_rst_Listas!HIPCIE_FLGREF) = 1 Then
            g_str_Parame = g_str_Parame & 6 & ", "
            g_str_Parame = g_str_Parame & 6 & ", "
         Else
            g_str_Parame = g_str_Parame & 4 & ", "
            g_str_Parame = g_str_Parame & 1 & ", "
         End If
      Else
         If Trim(g_rst_Listas!HIPCIE_FLGREF) = 1 Then
            g_str_Parame = g_str_Parame & 6 & ", "
            g_str_Parame = g_str_Parame & 6 & ", "
         Else
            g_str_Parame = g_str_Parame & Trim(g_rst_Listas!HIPCIE_SITCRE) & ", "
            g_str_Parame = g_str_Parame & Trim(g_rst_Listas!HIPCIE_SITCRE) & ", "
         End If
      End If
      
      'Tasa interes
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_TASINT, "###0.0000") & ", "
      
      'Capital aprobado
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_MTOPRE, "########0.00") & ", "
      
      'Capital desembolsado
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_MTOPRE, "########0.00") & ", "
      
      'Capital amortizado
      'If g_rst_Listas!HIPCIE_FLGREF = 1 Then
      '   g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_CAPVIG, "########0.00") & ", "
      'Else
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_CPTNAM + g_rst_Listas!HIPCIE_CPTCAM, "########0.00") & ", "
      'End If
      
      'Interes Devengado
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_ACUDVG, "########0.00") & ", "
      
      'Interes en suspenso
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_ACUDVC, "########0.00") & ", "
      
      'Intere moratorio
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_INTMOR, "########0.00") & ", "
      
      'Gastos de cobranza
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_GASCOB, "########0.00") & ", "
      
      'Garantia en moneda nacional y extranjera
      If g_rst_Listas!HIPCIE_TIPGAR = 1 Or g_rst_Listas!HIPCIE_TIPGAR = 2 Then
         If g_rst_Listas!HIPCIE_MONGAR = 1 Then
            g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_MTOGAR, "########0.00") & ", "
            g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         ElseIf g_rst_Listas!HIPCIE_MONGAR = 2 Then
            g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
            g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_MTOGAR, "########0.00") & ", "
         End If
      Else
         g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
      End If
      
      'Dias de mora
      g_str_Parame = g_str_Parame & Trim(g_rst_Listas!HIPCIE_DIAMOR) & ", "
      
      'Numero de cuotas atrasadas
      g_str_Parame = g_str_Parame & Trim(g_rst_Listas!HIPCIE_CUOATR) & ", "
      
      'Capital vencido
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_CAPVEN, "########0.00") & ", "
      
      'Ultima fecha de pago
      If g_rst_Listas!HIPCIE_ULTPAG <> 0 Then
         g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!HIPCIE_ULTPAG))) & "','DD/MM/YYYY'), "
      Else
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      'Ultima fecha de vencimiento del cronograma del cliente
      g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!HIPCIE_ULTVCT))) & "','DD/MM/YYYY'), "
      
      'Actividad economica
      g_str_Parame = g_str_Parame & Trim(g_rst_Listas!HIPCIE_ACTECO) & ", "
      
      'Saldo Capial del tramo no concesional
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_SALCAP, "########0.00") & ", "
      
      'Fecha de aprobacion del credito
      g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!HIPCIE_APRCRE))) & "','DD/MM/YYYY'), "
      
      'Interes capitalizado
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_INTCAP, "########0.00") & ", "
      
      'Capital de la ultima cuota pagada
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_UCPPAG, "########0.00") & ", "
      
      'Codigo CIIU
      g_str_Parame = g_str_Parame & Trim(g_rst_Listas!HIPCIE_CODCIU) & ", "
      
      'Fecha desembolso
      g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!HIPCIE_FECDES))) & "','DD/MM/YYYY'), "
      
      'Saldo Concesional
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_SALCON - g_rst_Listas!HIPCIE_PERPBP, "########0.00") & ", "
      
      'Tasa de Costo efectivo anual
      g_str_Parame = g_str_Parame & Format(g_rst_Listas!HIPCIE_COSEFE, "###0.0000") & ") "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
      
      '***************************************************************
      ' PARA CLIENTES CON CLASIFICACION ALINEADA 'DUDOSO' O 'PERDIDA'
      '***************************************************************
      If (g_rst_Listas!HIPCIE_CLAPRV = 3 Or g_rst_Listas!HIPCIE_CLAPRV = 4) And (g_rst_Listas!HIPCIE_CLACLI = 0) Then
         If g_rst_Listas!HIPCIE_FLGREF = 0 Then
            'ACTUALIZA PADRON
            modprc_g_str_CadEje = ""
            modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CREDITO_CIERRE_FINMES "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET INTERES = 0, "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       INTERES_COMP = " & Format(g_rst_Listas!HIPCIE_ACUDVG, "#####0.00") & " "
            modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE MES  = " & CStr(p_PerMes) & " "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND ANO  = " & CStr(p_PerAno) & " "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND CREDITO = '" & Trim(g_rst_Listas!HIPCIE_NUMOPE) & "'"
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, g_rst_Genera, 2) Then
               Exit Sub
            End If
            
            'ACTUALIZA MAESTO DE CIERRE
            modprc_g_str_CadEje = ""
            modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE CRE_HIPCIE "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET HIPCIE_ACUDVG = 0, "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "       HIPCIE_ACUDVC = " & Format(g_rst_Listas!HIPCIE_ACUDVG, "#####0.00") & " "
            modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE HIPCIE_PERMES = " & CStr(p_PerMes) & " "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCIE_PERANO = " & CStr(p_PerAno) & " "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND HIPCIE_NUMOPE = '" & Trim(g_rst_Listas!HIPCIE_NUMOPE) & "'"
         
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, g_rst_Genera, 2) Then
               Exit Sub
            End If
         End If
      End If
      
      'siguiente registro
      g_rst_Listas.MoveNext
   Loop
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub
Public Sub gs_Cargar_CreCom(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.*, B.* FROM TPR_CAFCIE A "
   g_str_Parame = g_str_Parame & "  INNER JOIN TPR_ETECIE B ON ETECIE_PERMES = CAFCIE_PERMES AND ETECIE_PERANO = CAFCIE_PERANO AND ETECIE_TIPDOC = CAFCIE_TIPDOC AND ETECIE_NUMDOC = CAFCIE_NUMDOC "
   g_str_Parame = g_str_Parame & "  WHERE CAFCIE_PERMES = " & p_PerMes & " AND CAFCIE_PERANO = " & p_PerAno & " AND CAFCIE_CODPRD = '008' "
   g_str_Parame = g_str_Parame & "  ORDER BY CAFCIE_NUMREF ASC "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      
      Do While Not g_rst_Listas.EOF
      
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO CREDITO_CIERRE_FINMES("
         g_str_Parame = g_str_Parame & "   CREDITO, "
         g_str_Parame = g_str_Parame & "   MES, "
         g_str_Parame = g_str_Parame & "   ANO, "
         g_str_Parame = g_str_Parame & "   COD_MONEDA, "
         g_str_Parame = g_str_Parame & "   CLIENTE, "
         g_str_Parame = g_str_Parame & "   PRODUCTO, "
         g_str_Parame = g_str_Parame & "   FLAG_ESTADO_CRED, "
         g_str_Parame = g_str_Parame & "   FLAG_ESTADO_ANT, "
         g_str_Parame = g_str_Parame & "   TASA, "
         g_str_Parame = g_str_Parame & "   CAPITAL_APROBADO, "
         g_str_Parame = g_str_Parame & "   CAPITAL_DESEMBOLSADO, "
         g_str_Parame = g_str_Parame & "   CAPITAL_AMORTIZADO, "
         g_str_Parame = g_str_Parame & "   INTERES, "
         g_str_Parame = g_str_Parame & "   INTERES_COMP, "
         g_str_Parame = g_str_Parame & "   INTERES_MOR, "
         g_str_Parame = g_str_Parame & "   OTROS_CARGOS, "
         g_str_Parame = g_str_Parame & "   SALDO_GARANTIA_MN, "
         g_str_Parame = g_str_Parame & "   SALDO_GARANTIA_ME, "
         g_str_Parame = g_str_Parame & "   DIAS_MOROSIDAD, "
         g_str_Parame = g_str_Parame & "   CUOTAS_ATRAZADAS, "
         g_str_Parame = g_str_Parame & "   CAPITAL_VENCIDO, "
         g_str_Parame = g_str_Parame & "   FECHA_ULT_MOV, "
         g_str_Parame = g_str_Parame & "   FECHA_VENCIMIENTO, "
         g_str_Parame = g_str_Parame & "   ACT_ECONOMICA, "
         g_str_Parame = g_str_Parame & "   SALDO_NO_CONCESIONAL, "
         g_str_Parame = g_str_Parame & "   FECHA_APROBACION, "
         g_str_Parame = g_str_Parame & "   CAPITAL_INTERES, "
         g_str_Parame = g_str_Parame & "   CAPITAL_AMORTIZADO_MES, "
         g_str_Parame = g_str_Parame & "   CIIU, "
         g_str_Parame = g_str_Parame & "   FECHA_DESEMBOLSO, "
         g_str_Parame = g_str_Parame & "   SALDO_CONCESIONAL, "
         g_str_Parame = g_str_Parame & "   TASA_COSTO_EFECTIVO) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!CAFCIE_NUMREF) & "', "
         g_str_Parame = g_str_Parame & p_PerMes & ", "
         g_str_Parame = g_str_Parame & p_PerAno & ", "
         g_str_Parame = g_str_Parame & "'" & Format(Trim(g_rst_Listas!CAFCIE_FIAMON), "000") & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!CAFCIE_NUMDOC) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!CAFCIE_CODPRD) & "', "
         
         If Trim(g_rst_Listas!CAFCIE_SITCRE) = 5 Then
            If Trim(g_rst_Listas!CAFCIE_FLGREF) = 1 Then
               g_str_Parame = g_str_Parame & 6 & ", "
               g_str_Parame = g_str_Parame & 6 & ", "
            Else
               g_str_Parame = g_str_Parame & 4 & ", "
               g_str_Parame = g_str_Parame & 1 & ", "
            End If
         Else
            If Trim(g_rst_Listas!CAFCIE_FLGREF) = 1 Then
               g_str_Parame = g_str_Parame & 6 & ", "
               g_str_Parame = g_str_Parame & 6 & ", "
            Else
               g_str_Parame = g_str_Parame & Trim(g_rst_Listas!CAFCIE_SITCRE) & ", "
               g_str_Parame = g_str_Parame & Trim(g_rst_Listas!CAFCIE_SITCRE) & ", "
            End If
         End If
         
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_FIATAS, "###0.0000") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_FIAIMP, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_FIAIMP, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_ACUDVG, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_ACUDVC, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_INTMOR, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_GASCOB, "########0.00") & ", "
         
'         If g_rst_Listas!CAFCIE_GARTIP = 1 Or g_rst_Listas!CAFCIE_GARTIP = 2 Then                           'CAFCIE_TIPGAR
'            If g_rst_Listas!CAFCIE_GARMON = 1 Then                                                          'CAFCIE_MONGAR
'               g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_GARMTO, "########0.00") & ", "      'CAFCIE_MTOGAR
'               g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
'            ElseIf g_rst_Listas!CAFCIE_GARMON = 2 Then
'               g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
'               g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_GARMTO, "########0.00") & ", "
'            End If
'         Else
'            g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
'            g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
'         End If
         
         If g_rst_Listas!ETECIE_GARHIP > 0 Then
            g_str_Parame = g_str_Parame & Format(g_rst_Listas!ETECIE_GARHIP, "########0.00") & ", "
            g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         Else
            g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
            g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         End If

         g_str_Parame = g_str_Parame & Trim(g_rst_Listas!CAFCIE_DIAMOR) & ", "
         g_str_Parame = g_str_Parame & Trim(g_rst_Listas!CAFCIE_CUOATR) & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_CAPVEN, "########0.00") & ", "
         
         If g_rst_Listas!CAFCIE_ULTPAG <> 0 Then
            g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!CAFCIE_ULTPAG))) & "','DD/MM/YYYY'), "
         Else
            g_str_Parame = g_str_Parame & "'', "
         End If
               
         g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!CAFCIE_ULTVCT))) & "','DD/MM/YYYY'), "
         g_str_Parame = g_str_Parame & "'" & "1.F" & "', "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_SALCAP, "########0.00") & ", "
         g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!CAFCIE_APRCRE))) & "','DD/MM/YYYY'), "
         g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_UCPPAG, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Trim(g_rst_Listas!ETECIE_CODCIU) & ", "
         g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!CAFCIE_FIAEMI))) & "','DD/MM/YYYY'), "
         g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!CAFCIE_COSEFE, "###0.0000") & ") "
                  
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub gs_Cargar_CreCom_OLD(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
   g_str_Parame = ""
   g_str_Parame = "SELECT * FROM CRE_COMCIE WHERE COMCIE_PERMES = " & p_PerMes & " AND COMCIE_PERANO = " & p_PerAno & " "
   g_str_Parame = g_str_Parame & "ORDER BY COMCIE_NUMOPE ASC "
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
      
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO CREDITO_CIERRE_FINMES("
         g_str_Parame = g_str_Parame & "   CREDITO, "
         g_str_Parame = g_str_Parame & "   MES, "
         g_str_Parame = g_str_Parame & "   ANO, "
         g_str_Parame = g_str_Parame & "   COD_MONEDA, "
         g_str_Parame = g_str_Parame & "   CLIENTE, "
         g_str_Parame = g_str_Parame & "   PRODUCTO, "
         g_str_Parame = g_str_Parame & "   FLAG_ESTADO_CRED, "
         g_str_Parame = g_str_Parame & "   FLAG_ESTADO_ANT, "
         g_str_Parame = g_str_Parame & "   TASA, "
         g_str_Parame = g_str_Parame & "   CAPITAL_APROBADO, "
         g_str_Parame = g_str_Parame & "   CAPITAL_DESEMBOLSADO, "
         g_str_Parame = g_str_Parame & "   CAPITAL_AMORTIZADO, "
         g_str_Parame = g_str_Parame & "   INTERES, "
         g_str_Parame = g_str_Parame & "   INTERES_COMP, "
         g_str_Parame = g_str_Parame & "   INTERES_MOR, "
         g_str_Parame = g_str_Parame & "   OTROS_CARGOS, "
         g_str_Parame = g_str_Parame & "   SALDO_GARANTIA_MN, "
         g_str_Parame = g_str_Parame & "   SALDO_GARANTIA_ME, "
         g_str_Parame = g_str_Parame & "   DIAS_MOROSIDAD, "
         g_str_Parame = g_str_Parame & "   CUOTAS_ATRAZADAS, "
         g_str_Parame = g_str_Parame & "   CAPITAL_VENCIDO, "
         g_str_Parame = g_str_Parame & "   FECHA_ULT_MOV, "
         g_str_Parame = g_str_Parame & "   FECHA_VENCIMIENTO, "
         g_str_Parame = g_str_Parame & "   ACT_ECONOMICA, "
         g_str_Parame = g_str_Parame & "   SALDO_NO_CONCESIONAL, "
         g_str_Parame = g_str_Parame & "   FECHA_APROBACION, "
         g_str_Parame = g_str_Parame & "   CAPITAL_INTERES, "
         g_str_Parame = g_str_Parame & "   CAPITAL_AMORTIZADO_MES, "
         g_str_Parame = g_str_Parame & "   CIIU, "
         g_str_Parame = g_str_Parame & "   FECHA_DESEMBOLSO, "
         g_str_Parame = g_str_Parame & "   SALDO_CONCESIONAL, "
         g_str_Parame = g_str_Parame & "   TASA_COSTO_EFECTIVO) "
         
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!COMCIE_NUMOPE) & "', "
         g_str_Parame = g_str_Parame & p_PerMes & ", "
         g_str_Parame = g_str_Parame & p_PerAno & ", "
         g_str_Parame = g_str_Parame & "'" & Format(Trim(g_rst_Listas!COMCIE_TIPMON), "000") & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!COMCIE_NDOCLI) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Listas!comcie_codprd) & "', "
         
         If Trim(g_rst_Listas!COMCIE_SITCRE) = 5 Then
            If Trim(g_rst_Listas!COMCIE_FLGREF) = 1 Then
               g_str_Parame = g_str_Parame & 6 & ", "
               g_str_Parame = g_str_Parame & 6 & ", "
            Else
               g_str_Parame = g_str_Parame & 4 & ", "
               g_str_Parame = g_str_Parame & 1 & ", "
            End If
         Else
            If Trim(g_rst_Listas!COMCIE_FLGREF) = 1 Then
               g_str_Parame = g_str_Parame & 6 & ", "
               g_str_Parame = g_str_Parame & 6 & ", "
            Else
               g_str_Parame = g_str_Parame & Trim(g_rst_Listas!COMCIE_SITCRE) & ", "
               g_str_Parame = g_str_Parame & Trim(g_rst_Listas!COMCIE_SITCRE) & ", "
            End If
         End If
         
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_TASINT, "###0.0000") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_MTOPRE, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_MTOPRE, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_ACUDVG, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_ACUDVC, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_INTMOR, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_GASCOB, "########0.00") & ", "
         
         If g_rst_Listas!comcie_tipgar = 1 Or g_rst_Listas!comcie_tipgar = 2 Then
            If g_rst_Listas!COMCIE_MONGAR = 1 Then
               g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_MTOGAR, "########0.00") & ", "
               g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
            ElseIf g_rst_Listas!COMCIE_MONGAR = 2 Then
               g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
               g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_MTOGAR, "########0.00") & ", "
            End If
         Else
            g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
            g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         End If
         
         g_str_Parame = g_str_Parame & Trim(g_rst_Listas!COMCIE_DIAMOR) & ", "
         g_str_Parame = g_str_Parame & Trim(g_rst_Listas!COMCIE_CUOATR) & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_CAPVEN, "########0.00") & ", "
         
         If g_rst_Listas!COMCIE_ULTPAG <> 0 Then
            g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!COMCIE_ULTPAG))) & "','DD/MM/YYYY'), "
         Else
            g_str_Parame = g_str_Parame & "'', "
         End If
               
         g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!COMCIE_ULTVCT))) & "','DD/MM/YYYY'), "
         g_str_Parame = g_str_Parame & "'" & "1.F" & "', "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_SALCAP, "########0.00") & ", "
         g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!COMCIE_APRCRE))) & "','DD/MM/YYYY'), "
         g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_UCPPAG, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Trim(g_rst_Listas!comcie_codciu) & ", "
         g_str_Parame = g_str_Parame & "to_date ('" & CDate(gf_FormatoFecha(Trim(g_rst_Listas!COMCIE_FECDES))) & "','DD/MM/YYYY'), "
         g_str_Parame = g_str_Parame & Format(0, "########0.00") & ", "
         g_str_Parame = g_str_Parame & Format(g_rst_Listas!COMCIE_COSEFE, "###0.0000") & ") "
                  
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Public Sub gs_Datos_CafCie(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TPR_CAFCIE ("
      g_str_Parame = g_str_Parame & "" & p_PerMes & ", "
      g_str_Parame = g_str_Parame & "" & p_PerAno & ", "
              
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
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
   Loop
End Sub

Public Sub gs_CafCie_VtaPat(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_TPR_VTAPAT_CIE ( "
      g_str_Parame = g_str_Parame & "" & p_PerAno & ", "
      g_str_Parame = g_str_Parame & "" & p_PerMes & ", "
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
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
   Loop
End Sub

Public Sub gs_Datos_BieAdj(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_BIEADJ_CIERRE ("
      g_str_Parame = g_str_Parame & "" & p_PerMes & ", "
      g_str_Parame = g_str_Parame & "" & p_PerAno & ", "
              
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
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
   Loop
End Sub
