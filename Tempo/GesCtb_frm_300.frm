VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Pro_CuoHip_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   3570
   ClientLeft      =   14595
   ClientTop       =   4770
   ClientWidth     =   7275
   Icon            =   "GesCtb_frm_300.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7395
      _Version        =   65536
      _ExtentX        =   13044
      _ExtentY        =   6588
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
            Left            =   630
            TabIndex        =   8
            Top             =   30
            Width           =   3465
            _Version        =   65536
            _ExtentX        =   6112
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Anexo 7 - 16 - 16b"
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
            Picture         =   "GesCtb_frm_300.frx":000C
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
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_300.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Procesar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6600
            Picture         =   "GesCtb_frm_300.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
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
         Begin VB.ComboBox cmb_Sucurs 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   450
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
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
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
         Begin VB.Label Label13 
            Caption         =   "Sucursal:"
            Height          =   255
            Left            =   90
            TabIndex        =   17
            Top             =   510
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   1200
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   90
            TabIndex        =   13
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   90
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   15
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
            TabIndex        =   16
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
Attribute VB_Name = "frm_Pro_CuoHip_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Empres()     As moddat_tpo_Genera
Dim l_arr_Sucurs()     As moddat_tpo_Genera
Dim r_str_Fechas(26)   As String
Dim l_str_CodEmp       As String
Dim l_str_CodSuc       As String
Dim l_str_PerMes       As String
Dim l_str_PerAno       As String
Dim l_lng_TotReg       As Long

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
      If MsgBox("¿Está seguro de Realizar el Proceso del Anexo 7/16/16B?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   Else
      If MsgBox("¿Está seguro de Reprocesar el Anexo 7/16/16B?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      g_str_Parame = "DELETE FROM RPT_ANEXOS WHERE "
      g_str_Parame = g_str_Parame & "ANEXOS_PERANO = " & l_str_PerAno & " AND "
      g_str_Parame = g_str_Parame & "ANEXOS_PERMES = " & l_str_PerMes & " AND "
      g_str_Parame = g_str_Parame & "ANEXOS_CODEMP = " & l_str_CodEmp & " AND "
      g_str_Parame = g_str_Parame & "ANEXOS_CODSUC = " & l_str_CodSuc & " "
                      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   End If
   
   cmd_Proces.Enabled = False
   Screen.MousePointer = 11
   pnl_BarPro.FloodPercent = 0
      
   'Totales
   Call ff_CanReg
   r_str_FeInEj = Format(Now, "YYYYMMDD")
   r_str_HoInEj = Format(Now, "HHMMSS")
   
   Call modprc_ctbp4001("CBRP4001", r_str_FeInEj, r_str_HoInEj, l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo, l_arr_Sucurs(cmb_Sucurs.ListIndex + 1).Genera_Codigo, l_str_PerMes, l_str_PerAno, l_lng_TotReg, pnl_BarPro)
   pnl_BarPro.FloodPercent = CDbl(Format(l_lng_TotReg / l_lng_TotReg * 100, "##0.00"))
   
   Screen.MousePointer = 0
   MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   cmd_Proces.Enabled = True
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   Call moddat_gs_Carga_LisIte_Combo(cmb_Period, 1, "033")
End Sub

Public Function ff_Buscar() As Integer
   ff_Buscar = 0
   g_str_Parame = "SELECT COUNT(*) AS TOTREG FROM RPT_ANEXOS WHERE "
   g_str_Parame = g_str_Parame & "ANEXOS_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "ANEXOS_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "ANEXOS_CODEMP = " & l_str_CodEmp & " AND "
   g_str_Parame = g_str_Parame & "ANEXOS_CODSUC = " & l_str_CodSuc & " "
   
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

Public Sub ff_CanReg()
   Dim l_int_mes           As Integer
   Dim l_int_ano           As Integer
   Dim l_str_fec           As String
   Dim l_str_aux           As String
   Dim l_dat_fec           As Date
   Dim l_dat_aux           As Date
   Dim l_int_con           As Integer
   Dim r_int_ConAux        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConTem        As Integer
   Dim r_int_AuxTem        As Integer
   
   Erase r_str_Fechas
      
   l_int_mes = l_str_PerMes
   l_int_ano = l_str_PerAno
   l_str_fec = "01/" & l_int_mes & "/" & l_int_ano
   l_dat_fec = modsec_gf_Fin_Del_Mes(CDate(l_str_fec))
   l_str_aux = modsec_gf_Fin_Del_Mes(CDate(l_str_fec))
   l_lng_TotReg = 0
   
   r_str_Fechas(0) = CDate(l_dat_fec + 1)
   r_str_Fechas(1) = CDate(l_dat_fec + 7)
   r_str_Fechas(2) = CDate(l_dat_fec + 8)
   r_str_Fechas(3) = CDate(l_dat_fec + 15)
   r_str_Fechas(4) = CDate(l_dat_fec + 16)
   l_str_fec = l_dat_fec + 1
   r_str_Fechas(5) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_str_fec, 5), 2), Right(l_str_fec, 4))))
   r_str_Fechas(6) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_str_fec, 5), 2), Right(l_str_fec, 4)))) + 1
   l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_str_fec, 5), 2), Right(l_str_fec, 4)))) + 1
   r_str_Fechas(7) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4)))) - 1
   r_str_Fechas(8) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
   l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
   r_str_Fechas(9) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4)))) - 1
   r_str_Fechas(10) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
   l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
   l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
   l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
   r_str_Fechas(11) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4)))) - 1
   r_str_Fechas(12) = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
   l_dat_fec = CDate(l_dat_fec + CInt(ff_Ultimo_Dia_Mes(Right(Left(l_dat_fec, 5), 2), Right(l_dat_fec, 4))))
   r_str_Fechas(13) = CDate(CDate(l_str_fec) + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
   r_str_Fechas(14) = CDate(CDate(l_str_fec) + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   l_dat_fec = CDate(CDate(l_str_fec) + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
      r_str_Fechas(15) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
   Else
      r_str_Fechas(15) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   End If
         
   If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
      r_str_Fechas(16) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   Else
      r_str_Fechas(16) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) + 1
   End If
   
   l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   r_str_Fechas(17) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
   r_str_Fechas(18) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   r_str_Fechas(19) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
   r_str_Fechas(20) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   r_str_Fechas(21) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
   r_str_Fechas(22) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   
   For l_int_con = 1 To 5
      l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   Next l_int_con
   
   If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
      r_str_Fechas(23) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
   Else
      r_str_Fechas(23) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   End If
         
   If l_int_mes = 1 Or l_int_mes = 2 Or l_int_mes = 12 Then
      r_str_Fechas(24) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   Else
      r_str_Fechas(24) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) + 1
   End If
   
   For l_int_con = 1 To 10
      l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   Next l_int_con
   
   r_str_Fechas(25) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4))))) - 1
   r_str_Fechas(26) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
         
'   For l_int_con = 1 To 20
'      l_dat_fec = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
'   Next l_int_con
'   .Cells(3, 15) = CDate(l_dat_fec + CInt(modsec_gf_Dias_del_Año(CInt(Right(l_dat_fec, 4)))))
   
   '**********************************************************************************************************************************************************
   
   'CUENTAS POR PAGAR
   For r_int_Contad = 0 To 26 Step 2
      g_str_Parame = ""
      g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
      If r_int_Contad = 26 Then
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            l_lng_TotReg = l_lng_TotReg + 1
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
     
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next r_int_Contad
   
   For r_int_Contad = 0 To 26 Step 2
      g_str_Parame = ""
      g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 2 AND "
      If r_int_Contad = 26 Then
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      'g_str_Parame = g_str_Parame & "(HIPMAE_NUMOPE <> '0040700001' AND HIPMAE_NUMOPE <> '0040700002' AND "
      'g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE <> '0040700003' AND HIPMAE_NUMOPE <> '0040700004') "
      g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            l_lng_TotReg = l_lng_TotReg + 1
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next r_int_Contad
      
   For r_int_Contad = 0 To 26 Step 2
      g_str_Parame = ""
      g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 4 AND "
      If r_int_Contad = 26 Then
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      'g_str_Parame = g_str_Parame & "(hipmae_numope = '0040700001' or hipmae_numope = '0040700002' or "
      'g_str_Parame = g_str_Parame & "hipmae_numope = '0040700003' or hipmae_numope = '0040700004') "
      g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            l_lng_TotReg = l_lng_TotReg + 1
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next r_int_Contad
   
   'CUENTAS POR COBRAR
   For r_int_Contad = 0 To 26 Step 2
      g_str_Parame = ""
      g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 3 AND "
      If r_int_Contad = 26 Then
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            l_lng_TotReg = l_lng_TotReg + 1
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
     
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next r_int_Contad
   
   For r_int_Contad = 0 To 26 Step 2
   
      g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 5 AND "
      If r_int_Contad = 26 Then
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            l_lng_TotReg = l_lng_TotReg + 1
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
     
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next r_int_Contad
   
   For r_int_Contad = 0 To 26 Step 2
   
      g_str_Parame = "SELECT HIPMAE_CODPRD, SUM(HIPCUO_CAPITA) AS CAPITAL FROM CRE_HIPMAE A, CRE_HIPCUO B WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = HIPMAE_NUMOPE AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 4 AND "
      If r_int_Contad = 26 Then
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " "
      Else
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT >= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad)) & " AND "
         g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & modsec_gf_GenFec(r_str_Fechas(r_int_Contad + 1)) & " "
      End If
      g_str_Parame = g_str_Parame & "GROUP BY HIPMAE_CODPRD "
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            l_lng_TotReg = l_lng_TotReg + 1
            g_rst_Princi.MoveNext
            DoEvents
         Loop
      End If
     
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next r_int_Contad

End Sub
