VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Pro_LimGlo_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   10080
   ClientTop       =   6795
   ClientWidth     =   7305
   Icon            =   "GesCtb_frm_904.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7335
      _Version        =   65536
      _ExtentX        =   12938
      _ExtentY        =   6429
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
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
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
            TabIndex        =   8
            Top             =   30
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Limites Globales"
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
            Height          =   375
            Left            =   630
            TabIndex        =   9
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
            Picture         =   "GesCtb_frm_904.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   10
         Top             =   780
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6600
            Picture         =   "GesCtb_frm_904.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_904.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Procesar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1575
         Left            =   30
         TabIndex        =   11
         Top             =   1470
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   2778
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   5025
         End
         Begin VB.ComboBox cmb_Period 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   810
            Width           =   2265
         End
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   450
            Width           =   5025
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   1170
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
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   90
            TabIndex        =   15
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   90
            TabIndex        =   14
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   1200
            Width           =   885
         End
         Begin VB.Label Label13 
            Caption         =   "Sucursal:"
            Height          =   255
            Left            =   90
            TabIndex        =   12
            Top             =   510
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   16
         Top             =   3090
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
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
            TabIndex        =   17
            Top             =   60
            Width           =   7095
            _Version        =   65536
            _ExtentX        =   12515
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
Attribute VB_Name = "frm_Pro_LimGlo_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Empres()     As moddat_tpo_Genera
Dim l_arr_Sucurs()     As moddat_tpo_Genera
Dim l_str_CodEmp       As String
Dim l_str_CodSuc       As String
Dim l_str_PerMes       As String
Dim l_str_PerAno       As String
Dim l_lng_TotReg       As Long
Dim l_lng_NumReg       As Long

Private Sub cmb_Empres_Click()
   If cmb_Empres.ListIndex > -1 Then
      Call moddat_gs_Carga_SucAge(cmb_Sucurs, l_arr_Sucurs, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo)
      Call gs_SetFocus(cmb_Sucurs)
   Else
      cmb_Sucurs.Clear
   End If
End Sub

Private Sub cmb_Sucurs_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Period)
   End If
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Empres)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Proces_Click()
   Dim r_str_FeInEj  As String
   Dim r_str_HoInEj  As String
   
   If cmb_Empres.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Empres)
      Exit Sub
   End If
         
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
   
   l_str_PerAno = Format(ipp_PerAno.Text, "0000")
   l_str_PerMes = Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00")
   l_str_CodEmp = l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo
   l_str_CodSuc = l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo
      
   If ff_Buscar = 0 Then
      If MsgBox("¿Está seguro de Realizar el Proceso de Limites Globales?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   Else
      If MsgBox("¿Está seguro de Reprocesar Limites Globales?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      g_str_Parame = "DELETE FROM TMP_LIMGLO WHERE "
      g_str_Parame = g_str_Parame & "LIMGLO_PERANO = " & l_str_PerAno & " AND "
      g_str_Parame = g_str_Parame & "LIMGLO_PERMES = " & l_str_PerMes & " AND "
      g_str_Parame = g_str_Parame & "LIMGLO_CODEMP = " & l_str_CodEmp & " AND "
      g_str_Parame = g_str_Parame & "LIMGLO_CODSUC = " & l_str_CodSuc & " "
                      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   End If
   
   cmd_Proces.Enabled = False
   Screen.MousePointer = 11
   
   'Totales
   l_lng_NumReg = 0
   l_lng_TotReg = ff_TotHip + ff_TotCom
   r_str_FeInEj = Format(Now, "YYYYMMDD")
   r_str_HoInEj = Format(Now, "HHMMSS")
   
   'Call modprc_ctbp4001("CBRP4001", r_str_FeInEj, r_str_HoInEj, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo, l_str_PerMes, l_str_PerAno, l_lng_TotReg, pnl_BarPro)
   Call fs_SalHip(pnl_BarPro)
   Call fs_SalCom(pnl_BarPro)
   
   pnl_BarPro.FloodPercent = CDbl(Format(l_lng_TotReg / l_lng_TotReg * 100, "##0.00"))
      
   MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   Screen.MousePointer = 0
   cmd_Proces.Enabled = True
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   Call moddat_gs_Carga_LisIte_Combo(cmb_Period, 1, "033")
End Sub

Private Sub fs_SalHip(Optional p_BarPro As SSPanel)
   Dim r_int_FecVct     As Double
   Dim r_dbl_TncMen_01  As Double
   Dim r_dbl_TncMay_01  As Double
   Dim r_dbl_TcMen_01   As Double
   Dim r_dbl_TcMay_01   As Double
   Dim r_dbl_SalMen     As Double
   Dim r_dbl_SalMay     As Double
   Dim r_dbl_MtoMen     As Double
   Dim r_dbl_MtoMay     As Double
   
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'Proceso
   Screen.MousePointer = 11
   p_BarPro.FloodPercent = 0
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = "SELECT * FROM CRE_HIPMAE, CRE_HIPCIE, CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = HIPMAE_TDOCLI AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = HIPMAE_NDOCLI AND "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = HIPCIE_NUMOPE AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_int_FecVct = ff_FecVct(g_rst_Princi!HIPMAE_NUMOPE, Trim(g_rst_Princi!HIPMAE_NUMCUO - g_rst_Princi!HIPMAE_CUOPEN))
         r_dbl_TncMen_01 = ff_TncMen_01(g_rst_Princi!HIPMAE_NUMOPE, r_int_FecVct)
         r_dbl_TncMay_01 = ff_TncMay_01(g_rst_Princi!HIPMAE_NUMOPE, Trim(g_rst_Princi!HIPMAE_NUMCUO - g_rst_Princi!HIPMAE_CUOPEN))
         r_dbl_TcMen_01 = ff_TcMen_01(g_rst_Princi!HIPMAE_NUMOPE, r_int_FecVct)
         r_dbl_TcMay_01 = ff_TcMay_01(g_rst_Princi!HIPMAE_NUMOPE, r_int_FecVct)
   
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO TMP_LIMGLO("
         g_str_Parame = g_str_Parame & "LIMGLO_CODEMP, "
         g_str_Parame = g_str_Parame & "LIMGLO_CODSUC, "
         g_str_Parame = g_str_Parame & "LIMGLO_PERMES, "
         g_str_Parame = g_str_Parame & "LIMGLO_PERANO, "
         g_str_Parame = g_str_Parame & "LIMGLO_TDOCLI, "
         g_str_Parame = g_str_Parame & "LIMGLO_NDOCLI, "
         g_str_Parame = g_str_Parame & "LIMGLO_NOMBRE, "
         g_str_Parame = g_str_Parame & "LIMGLO_TIPGAR, "
         g_str_Parame = g_str_Parame & "LIMGLO_MONEDA, "
         g_str_Parame = g_str_Parame & "LIMGLO_TNCMEN, "
         g_str_Parame = g_str_Parame & "LIMGLO_TNCMAY, "
         g_str_Parame = g_str_Parame & "LIMGLO_TCOMEN, "
         g_str_Parame = g_str_Parame & "LIMGLO_TCOMAY, "
         g_str_Parame = g_str_Parame & "LIMGLO_SALCAP, "
         g_str_Parame = g_str_Parame & "LIMGLO_SALCON, "
         g_str_Parame = g_str_Parame & "LIMGLO_SALMAY, "
         g_str_Parame = g_str_Parame & "LIMGLO_SALMEN, "
         g_str_Parame = g_str_Parame & "LIMGLO_SALDOL, "
         g_str_Parame = g_str_Parame & "LIMGLO_SALSOL) "
                  
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_CodSuc & "', "
         g_str_Parame = g_str_Parame & l_str_PerMes & ", "
         g_str_Parame = g_str_Parame & l_str_PerAno & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!DATGEN_TIPDOC & ", "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DATGEN_NUMDOC & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!DATGEN_APEPAT) + " " + Trim(g_rst_Princi!DATGEN_APEMAT) + " " + Trim(g_rst_Princi!DATGEN_NOMBRE) & "', "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_TIPGAR & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!hipmae_moneda & ", "
         g_str_Parame = g_str_Parame & r_dbl_TncMen_01 & ", "
         g_str_Parame = g_str_Parame & r_dbl_TncMay_01 & ", "
         g_str_Parame = g_str_Parame & r_dbl_TcMen_01 & ", "
         g_str_Parame = g_str_Parame & r_dbl_TcMay_01 & ", "
         g_str_Parame = g_str_Parame & r_dbl_TncMen_01 + r_dbl_TncMay_01 & ", "
         g_str_Parame = g_str_Parame & r_dbl_TcMen_01 + r_dbl_TcMay_01 & ", "
                             
         If g_rst_Princi!hipmae_moneda = 1 Then
            g_str_Parame = g_str_Parame & r_dbl_TncMay_01 + r_dbl_TcMay_01 & ", "
            g_str_Parame = g_str_Parame & r_dbl_TncMen_01 + r_dbl_TcMen_01 & ", "
            
         ElseIf g_rst_Princi!hipmae_moneda = 2 Then
            g_str_Parame = g_str_Parame & Format((r_dbl_TncMay_01 + r_dbl_TcMay_01) * g_rst_Princi!HIPCIE_TIPCAM, "###########0.00") & ", "
            g_str_Parame = g_str_Parame & Format((r_dbl_TncMen_01 + r_dbl_TcMen_01) * g_rst_Princi!HIPCIE_TIPCAM, "###########0.00") & ", "
         End If
                           
         If g_rst_Princi!hipmae_moneda = 1 Then
            g_str_Parame = g_str_Parame & Format(((g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON) / g_rst_Princi!HIPCIE_TIPCAM), "###########0.00") & ", "
         ElseIf g_rst_Princi!hipmae_moneda = 2 Then
            g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON & ", "
         End If
         
         If g_rst_Princi!hipmae_moneda = 1 Then
            g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON & ") "
         ElseIf g_rst_Princi!hipmae_moneda = 2 Then
            g_str_Parame = g_str_Parame & Format(((g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON) * g_rst_Princi!HIPCIE_TIPCAM), "###########0.00") & ") "
         End If
                         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
            
         g_rst_Princi.MoveNext
         l_lng_NumReg = l_lng_NumReg + 1
         DoEvents
         p_BarPro.FloodPercent = CDbl(Format(l_lng_NumReg / l_lng_TotReg * 100, "##0.00"))
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Screen.MousePointer = 0
      MsgBox "No se encontraron registros en el Proceso.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Sub fs_SalCom(Optional p_BarPro As SSPanel)
      
   'Proceso
   Screen.MousePointer = 11
   p_BarPro.FloodPercent = 0
      
   'Leyendo Tabla de solicitudes
   g_str_Parame = "SELECT COMCIE_NDOCLI, COMCIE_TDOCLI, SUM(COMCIE_MTOPRE) AS MTOPRE, MAX(COMCIE_TIPGAR) AS TIPGAR, MAX(COMCIE_CODPRD) AS CODPRD, MAX(COMCIE_TIPMON) AS TIPMON, MAX(COMCIE_TIPCAM) AS TIPCAM, MAX(DATGEN_RAZSOC) AS RAZSOC FROM CRE_COMCIE, EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_CODEMP = " & l_str_CodEmp & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_NDOCLI = DATGEN_EMPNDO "
   g_str_Parame = g_str_Parame & "GROUP BY COMCIE_NDOCLI, COMCIE_TDOCLI"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO TMP_LIMGLO("
         g_str_Parame = g_str_Parame & "LIMGLO_CODEMP, "
         g_str_Parame = g_str_Parame & "LIMGLO_CODSUC, "
         g_str_Parame = g_str_Parame & "LIMGLO_PERMES, "
         g_str_Parame = g_str_Parame & "LIMGLO_PERANO, "
         g_str_Parame = g_str_Parame & "LIMGLO_TDOCLI, "
         g_str_Parame = g_str_Parame & "LIMGLO_NDOCLI, "
         g_str_Parame = g_str_Parame & "LIMGLO_NOMBRE, "
         g_str_Parame = g_str_Parame & "LIMGLO_TIPGAR, "
         g_str_Parame = g_str_Parame & "LIMGLO_MONEDA, "
         g_str_Parame = g_str_Parame & "LIMGLO_SALDOL, "
         g_str_Parame = g_str_Parame & "LIMGLO_SALSOL) "
                  
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_CodSuc & "', "
         g_str_Parame = g_str_Parame & l_str_PerMes & ", "
         g_str_Parame = g_str_Parame & l_str_PerAno & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!COMCIE_TDOCLI & ", "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!COMCIE_NDOCLI) & "', "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!RAZSOC) & "', "
         g_str_Parame = g_str_Parame & g_rst_Princi!TIPGAR & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!TIPMON & ", "
                           
         If g_rst_Princi!TIPMON = 1 Then
            g_str_Parame = g_str_Parame & Format(g_rst_Princi!MTOPRE / g_rst_Princi!TIPCAM, "###########0.00") & ", "
         ElseIf g_rst_Princi!TIPMON = 2 Then
            g_str_Parame = g_str_Parame & g_rst_Princi!MTOPRE & ", "
         End If
         
         If g_rst_Princi!TIPMON = 1 Then
            g_str_Parame = g_str_Parame & g_rst_Princi!MTOPRE & ") "
         ElseIf g_rst_Princi!TIPMON = 2 Then
            g_str_Parame = g_str_Parame & Format(g_rst_Princi!MTOPRE * g_rst_Princi!TIPCAM, "###########0.00") & ") "
         End If
                         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
         g_rst_Princi.MoveNext
         l_lng_NumReg = l_lng_NumReg + 1
         DoEvents

         p_BarPro.FloodPercent = CDbl(Format(l_lng_NumReg / l_lng_TotReg * 100, "##0.00"))
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Screen.MousePointer = 0
      MsgBox "No se encontraron registros en el Proceso.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Function ff_FecVct(ByVal p_NumSol As String, ByVal l_int_NumCuo As Integer) As Double
   ff_FecVct = 0
   
   g_str_Parame = "select * from cre_hipcuo where "
   g_str_Parame = g_str_Parame & "hipcuo_numope = '" & p_NumSol & "' and "
   g_str_Parame = g_str_Parame & "hipcuo_tipcro = 1 and "
   g_str_Parame = g_str_Parame & "hipcuo_numcuo = " & l_int_NumCuo + 12

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ff_FecVct = g_rst_Listas!HIPCUO_FECVCT
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function ff_TncMen_01(ByVal p_NumSol As String, ByVal l_int_FecVct As Double) As Double
   ff_TncMen_01 = 0
   
   g_str_Parame = "select * from cre_hipcuo where "
   g_str_Parame = g_str_Parame & "hipcuo_numope = '" & p_NumSol & "' and "
   g_str_Parame = g_str_Parame & "hipcuo_tipcro = 1 and "
   g_str_Parame = g_str_Parame & "hipcuo_fecvct <= " & l_int_FecVct & " and "
   g_str_Parame = g_str_Parame & "hipcuo_situac = 2 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ff_TncMen_01 = ff_TncMen_01 + g_rst_Listas!HIPCUO_CAPITA
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function ff_TncMay_01(ByVal p_NumSol As String, ByVal l_int_NumCuo As Integer) As Double
   ff_TncMay_01 = 0
   
   g_str_Parame = "select * from cre_hipcuo where "
   g_str_Parame = g_str_Parame & "hipcuo_numope = '" & p_NumSol & "' and "
   g_str_Parame = g_str_Parame & "hipcuo_tipcro = 1 and "
   g_str_Parame = g_str_Parame & "hipcuo_numcuo > " & l_int_NumCuo + 12
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ff_TncMay_01 = ff_TncMay_01 + g_rst_Listas!HIPCUO_CAPITA
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function ff_TcMen_01(ByVal p_NumSol As String, ByVal l_int_FecVct As Double) As Double
   ff_TcMen_01 = 0
         
   g_str_Parame = "select * from cre_hipcuo where "
   g_str_Parame = g_str_Parame & "hipcuo_numope = '" & p_NumSol & "' and "
   
   'If p_NumSol = "0040700001" Or p_NumSol = "0040700002" Or p_NumSol = "0040700003" Or p_NumSol = "0040700004" Then
   '   g_str_Parame = g_str_Parame & "hipcuo_tipcro = 4 and "
   'Else
      g_str_Parame = g_str_Parame & "hipcuo_tipcro = 2 and "
   'End If
   
   g_str_Parame = g_str_Parame & "hipcuo_fecvct <= " & l_int_FecVct & " and "
   g_str_Parame = g_str_Parame & "hipcuo_situac = 2"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ff_TcMen_01 = ff_TcMen_01 + g_rst_Listas!HIPCUO_CAPITA
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function ff_TcMay_01(ByVal p_NumSol As String, ByVal l_int_FecVct As Double) As Double
   ff_TcMay_01 = 0
   
   g_str_Parame = "select * from cre_hipcuo where "
   g_str_Parame = g_str_Parame & "hipcuo_numope = '" & p_NumSol & "' and "
   
   'If p_NumSol = "0040700001" Or p_NumSol = "0040700002" Or p_NumSol = "0040700003" Or p_NumSol = "0040700004" Then
   '   g_str_Parame = g_str_Parame & "hipcuo_tipcro = 4 and "
   'Else
      g_str_Parame = g_str_Parame & "hipcuo_tipcro = 2 and "
   'End If
   
   g_str_Parame = g_str_Parame & "hipcuo_fecvct > " & l_int_FecVct
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ff_TcMay_01 = ff_TcMay_01 + g_rst_Listas!HIPCUO_CAPITA
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function ff_Buscar() As Integer
   ff_Buscar = 0
      
   g_str_Parame = "SELECT COUNT(*) AS TOTREG FROM TMP_LIMGLO WHERE "
   g_str_Parame = g_str_Parame & "LIMGLO_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "LIMGLO_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "LIMGLO_CODEMP = " & l_str_CodEmp & " AND "
   g_str_Parame = g_str_Parame & "LIMGLO_CODSUC = " & l_str_CodSuc & " "
                   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         ff_Buscar = g_rst_Princi!TOTREG
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Public Function ff_TotHip() As Integer
   ff_TotHip = 0

   g_str_Parame = "SELECT COUNT(*) AS TOTREG FROM CRE_HIPMAE, CRE_HIPCIE, CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = HIPMAE_TDOCLI AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = HIPMAE_NDOCLI AND "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = HIPCIE_NUMOPE AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         ff_TotHip = g_rst_Princi!TOTREG
         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Public Function ff_TotCom() As Integer
   ff_TotCom = 0

   g_str_Parame = "SELECT COMCIE_NDOCLI, COMCIE_TDOCLI, SUM(COMCIE_MTOPRE) AS MTOPRE, MAX(COMCIE_TIPGAR) AS TIPGAR, MAX(COMCIE_CODPRD) AS CODPRD, MAX(COMCIE_TIPMON) AS TIPMON, MAX(COMCIE_TIPCAM) AS TIPCAM, MAX(DATGEN_RAZSOC) AS RAZSOC FROM CRE_COMCIE, EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "COMCIE_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_CODEMP = " & l_str_CodEmp & " AND "
   g_str_Parame = g_str_Parame & "COMCIE_NDOCLI = DATGEN_EMPNDO "
   g_str_Parame = g_str_Parame & "GROUP BY COMCIE_NDOCLI, COMCIE_TDOCLI"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         ff_TotCom = ff_TotCom + 1
         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function
