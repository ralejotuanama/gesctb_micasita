VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Pro_BalCom_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   3510
   ClientLeft      =   5850
   ClientTop       =   4185
   ClientWidth     =   7275
   Icon            =   "GesCtb_frm_837.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7395
      _Version        =   65536
      _ExtentX        =   13044
      _ExtentY        =   6376
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
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Balance de Comprobación"
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
            Left            =   660
            TabIndex        =   15
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
            Picture         =   "GesCtb_frm_837.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   9
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
            Picture         =   "GesCtb_frm_837.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_837.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Procesar"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1515
         Left            =   30
         TabIndex        =   10
         Top             =   1470
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   2672
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
         Begin VB.CheckBox chk_CieEje 
            Caption         =   "Cierre del Ejercicio"
            Height          =   285
            Left            =   1560
            TabIndex        =   2
            Top             =   810
            Width           =   1995
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   5595
         End
         Begin VB.ComboBox cmb_Period 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   450
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   1110
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
            TabIndex        =   16
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   450
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   11
            Top             =   1140
            Width           =   885
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   13
         Top             =   3030
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
            TabIndex        =   14
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
Attribute VB_Name = "frm_Pro_BalCom_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_FecImp        As String
Dim l_str_HorImp        As String
Dim l_str_PerAno        As String
Dim l_str_PerMes        As String
Dim l_str_CodEmp        As String
Dim l_lng_NumReg        As Long
Dim l_lng_TotReg        As Long
Dim l_arr_Empres()      As moddat_tpo_Genera

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Period)
   End If
End Sub

Private Sub cmb_Period_Click()
   If (cmb_Period.ItemData(cmb_Period.ListIndex)) <> 12 Then
      chk_CieEje.Enabled = False
   Else
      chk_CieEje.Enabled = True
   End If
   chk_CieEje.Value = 0
End Sub

Private Sub cmb_Period_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(chk_CieEje)
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Empres)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_Period.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_Period, 1, "033")
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   ipp_PerAno = Mid(date, 7, 4)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_Proces_Click()
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

   If (cmb_Period.ItemData(cmb_Period.ListIndex) = 12) And (chk_CieEje.Value = 1) Then
      l_str_PerMes = Format(cmb_Period.ItemData(cmb_Period.ListIndex) + 1, "00")
   Else
      l_str_PerMes = Format(cmb_Period.ItemData(cmb_Period.ListIndex), "00")
   End If

   l_str_CodEmp = l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo

   If ff_Buscar = 0 Then
      If MsgBox("¿Está seguro de Realizar el Proceso de Balance de Comprobacion?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If

   Else
      If MsgBox("¿Está seguro de Reprocesar el Balance de Comprobacion?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If

      g_str_Parame = "DELETE FROM CTB_BALCOM WHERE "
      g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   End If

   cmd_Proces.Enabled = False
   Screen.MousePointer = 11

   'Totales
   l_lng_NumReg = 0
   l_lng_TotReg = ff_TotMon_Sol1 + ff_TotMon_Sol2 + ff_TotMon_Dol1 + ff_TotMon_Dol2
   l_lng_TotReg = l_lng_TotReg + (l_lng_TotReg * 0.1)

   'Cuentas Generales
   Call fs_BalCom_MonSol(pnl_BarPro)
   Call fs_BalCom_MonDol(pnl_BarPro)

   'Cuenta Integradoras
   Call fs_CueInt(1)
   Call fs_CueInt(2)

   'Saldo Anterior
   If l_str_PerMes <> 1 Then
      Call fs_SalAnt(1)
      Call fs_SalAnt(2)
   Else
      Call fs_SalEne_MonSol(1)
      Call fs_SalEne_MonDol(2)
   End If
   
   'Saldo Final
   Call fs_SalFin(1)
   Call fs_SalFin(2)
   
   'Depuracion de Cuentas
   Call fs_DepSal
        
   'Cuentas Hijos/Padres Soles
   Call fs_CueHij_MonSol
   Call fs_CueHij_MonDol
   
   'Cuentas Padres
   Call fs_CuePad(1)
   Call fs_CuePad(2)
         
   'Cuentas Restantes 101 - 102
   Call fs_CueFal(1)
   Call fs_CueFal(2)
   
   pnl_BarPro.FloodPercent = CDbl(Format(l_lng_TotReg / l_lng_TotReg * 100, "##0.00"))
   MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt

   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
   cmd_Proces.Enabled = True
End Sub

Private Sub fs_CueHij_MonSol()

   Dim r_dbl_ImpoMN        As Double
   Dim r_dbl_ImpoME        As Double

   Dim r_str_CodCta        As String
   Dim r_str_CtaAux        As String

   Dim r_int_FlagDH        As Integer
   Dim r_int_ConAux        As Integer
   Dim r_int_ConCue        As Integer
   Dim r_int_NroIte        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConVar        As Integer
   Dim r_int_Contad2       As Integer

   Dim r_dbl_MtoDeb        As Double
   Dim r_dbl_MtoHab        As Double

   Dim r_int_CadAx1        As Integer
   Dim r_int_CadAx2        As Integer
   Dim r_int_Iterac        As Integer

   Dim r_str_Cuenta        As String
   Dim r_str_DesCue        As String

   Dim r_Int_NroAux        As Integer

   Dim r_dbl_TipCam        As Double

   Dim r_arr_CtaCtb()      As modtac_tpo_CtaCtb
   Dim r_arr_CtaAux()      As modtac_tpo_CtaAux
   Dim r_arr_BalCom()      As modtac_tpo_BalCom

   ReDim r_arr_CtaCtb(0)
   ReDim r_arr_CtaAux(0)
   ReDim r_arr_BalCom(0)

   l_str_FecImp = Format(date, "yyyymmdd")
   l_str_HorImp = Format(Time, "hhmmss")

   'Para leer Cuentas Contables Agrupadas/Obtenemos DEBE-HABER y obtenemos los montos respectivos
   g_str_Parame = "SELECT * FROM CTB_BALCOM WHERE "
   g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & IIf(l_str_PerMes = 13, l_str_PerMes - 1, l_str_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_TIPMON = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY BALCOM_CUENTA ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF

         r_str_CtaAux = Trim(g_rst_Princi!BALCOM_CUENTA)
         r_int_CadAx1 = Len(Trim(g_rst_Princi!BALCOM_CUENTA))
         
         r_int_CadAx2 = 0
         
         'ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)
         
         'r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_CtaAux = r_str_CtaAux
         'r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpDeb = g_rst_Princi!BalCom_ImpDeb
         'r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpHab = g_rst_Princi!BalCom_ImpHab
         
         Do While r_int_CadAx1 > 4
               
            ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)
            
            r_int_CadAx2 = r_int_CadAx2 + 2
            r_int_CadAx1 = r_int_CadAx1 - 2
            
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_CtaAux = Mid(r_str_CtaAux, 1, r_int_CadAx1) & String(r_int_CadAx2, "0")
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpDeb = g_rst_Princi!BalCom_ImpDeb
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpHab = g_rst_Princi!BalCom_ImpHab
            
         Loop
         
         If r_int_CadAx1 = 4 Then
         
            ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)
         
            r_int_CadAx2 = r_int_CadAx2 + 1
            r_int_CadAx1 = r_int_CadAx1 - 1
            
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_CtaAux = Mid(r_str_CtaAux, 1, r_int_CadAx1) & String(r_int_CadAx2, "0")
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpDeb = g_rst_Princi!BalCom_ImpDeb
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpHab = g_rst_Princi!BalCom_ImpHab
                        
         End If
         
         g_rst_Princi.MoveNext
         
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   r_Int_NroAux = 1
      
   'Insertando Cuentas a la tabla CTB_BALCOM
   For r_int_NroIte = 1 To UBound(r_arr_CtaAux)
      
      If UBound(r_arr_BalCom) = 0 Then
      
         ReDim Preserve r_arr_BalCom(UBound(r_arr_BalCom) + 1)
         
         r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_CtaAux = r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux
         r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_ImpDeb = r_arr_CtaAux(r_int_NroIte).CtaAux_ImpDeb
         r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_ImpHab = r_arr_CtaAux(r_int_NroIte).CtaAux_ImpHab
         
      Else
      
         Do While r_Int_NroAux <= UBound(r_arr_BalCom)
            
         If CStr(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux)) = CStr(Trim(r_arr_BalCom(r_Int_NroAux).BalCom_CtaAux)) Then
            
            r_arr_BalCom(r_Int_NroAux).BalCom_ImpDeb = r_arr_BalCom(r_Int_NroAux).BalCom_ImpDeb + r_arr_CtaAux(r_int_NroIte).CtaAux_ImpDeb
            r_arr_BalCom(r_Int_NroAux).BalCom_ImpHab = r_arr_BalCom(r_Int_NroAux).BalCom_ImpHab + r_arr_CtaAux(r_int_NroIte).CtaAux_ImpHab
            
            Exit Do
         
         Else
            
            If UBound(r_arr_BalCom) = r_Int_NroAux Then
                              
               ReDim Preserve r_arr_BalCom(UBound(r_arr_BalCom) + 1)
                              
               r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_CtaAux = r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux
               r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_ImpDeb = r_arr_CtaAux(r_int_NroIte).CtaAux_ImpDeb
               r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_ImpHab = r_arr_CtaAux(r_int_NroIte).CtaAux_ImpHab
               
               Exit Do
            
            End If
            
         End If
                        
         r_Int_NroAux = r_Int_NroAux + 1
         
         Loop
                           
      End If
      
      r_Int_NroAux = 1
     
   Next r_int_NroIte
   
   
   'Insertando Cuenta Integradora a la Tabla CTB_BALCOM
   For r_int_NroIte = 1 To UBound(r_arr_BalCom) - 1
           
      r_str_DesCue = ff_DesCue(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux))
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CTB_BALCOM ("
      g_str_Parame = g_str_Parame & "BALCOM_CODEMP, "
      g_str_Parame = g_str_Parame & "BALCOM_PERMES, "
      g_str_Parame = g_str_Parame & "BALCOM_PERANO, "
      g_str_Parame = g_str_Parame & "BALCOM_TIPMON, "
      g_str_Parame = g_str_Parame & "BALCOM_TIPBAL, "
      g_str_Parame = g_str_Parame & "BALCOM_CUENTA, "
      g_str_Parame = g_str_Parame & "BALCOM_DESCRI, "
      g_str_Parame = g_str_Parame & "BALCOM_SLINDB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLINHB, "
      g_str_Parame = g_str_Parame & "BALCOM_IMPDEB, "
      g_str_Parame = g_str_Parame & "BALCOM_IMPHAB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLFIDB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLFIHB, "
      g_str_Parame = g_str_Parame & "BALCOM_FECCRE, "
      g_str_Parame = g_str_Parame & "BALCOM_USUCRE) "
                  
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_PerMes & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_PerAno & "', "
      g_str_Parame = g_str_Parame & 1 & ", "
      
      If Mid(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 3, 1) = 2 Then
         If (Left(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 1) = 4 Or Left(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 1) = 5) Then
            g_str_Parame = g_str_Parame & 1 & ", "
         Else
            g_str_Parame = g_str_Parame & 3 & ", "
         End If
      Else
         g_str_Parame = g_str_Parame & 0 & ", "
      End If
      
      g_str_Parame = g_str_Parame & "'" & r_arr_BalCom(r_int_NroIte).BalCom_CtaAux & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_DesCue & "', "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb & ", "
      g_str_Parame = g_str_Parame & r_arr_BalCom(r_int_NroIte).BalCom_ImpHab & ", "
      
      If (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb - r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) > 0 Then
         g_str_Parame = g_str_Parame & (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb - r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) & ", "
      Else
         g_str_Parame = g_str_Parame & 0 & ", "
      End If
      
      If (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb - r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) > 0 Then
         g_str_Parame = g_str_Parame & 0 & ", "
      Else
         If (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb) < (r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) Then
            g_str_Parame = g_str_Parame & (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb - r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) * -1 & ", "
         Else
            g_str_Parame = g_str_Parame & (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb - r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) & ", "
         End If
      End If
      
      g_str_Parame = g_str_Parame & "'" & l_str_FecImp & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
      
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
   
   Next r_int_NroIte
      

End Sub


Private Sub fs_CueHij_MonDol()

   Dim r_dbl_ImpoMN        As Double
   Dim r_dbl_ImpoME        As Double

   Dim r_str_CodCta        As String
   Dim r_str_CtaAux        As String

   Dim r_int_FlagDH        As Integer
   Dim r_int_ConAux        As Integer
   Dim r_int_ConCue        As Integer
   Dim r_int_NroIte        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConVar        As Integer
   Dim r_int_Contad2       As Integer

   Dim r_dbl_MtoDeb        As Double
   Dim r_dbl_MtoHab        As Double

   Dim r_int_CadAx1        As Integer
   Dim r_int_CadAx2        As Integer
   Dim r_int_Iterac        As Integer

   Dim r_str_Cuenta        As String
   Dim r_str_DesCue        As String

   Dim r_Int_NroAux        As Integer

   Dim r_dbl_TipCam        As Double

   Dim r_arr_CtaCtb()      As modtac_tpo_CtaCtb
   Dim r_arr_CtaAux()      As modtac_tpo_CtaAux
   Dim r_arr_BalCom()      As modtac_tpo_BalCom

   ReDim r_arr_CtaCtb(0)
   ReDim r_arr_CtaAux(0)
   ReDim r_arr_BalCom(0)

   l_str_FecImp = Format(date, "yyyymmdd")
   l_str_HorImp = Format(Time, "hhmmss")

   

   'Para leer Cuentas Contables Agrupadas/Obtenemos DEBE-HABER y obtenemos los montos respectivos
   g_str_Parame = "SELECT * FROM CTB_BALCOM WHERE "
   g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & IIf(l_str_PerMes = 13, l_str_PerMes - 1, l_str_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_TIPMON = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY BALCOM_CUENTA ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF

         r_str_CtaAux = Trim(g_rst_Princi!BALCOM_CUENTA)
         r_int_CadAx1 = Len(Trim(g_rst_Princi!BALCOM_CUENTA))
         
         r_int_CadAx2 = 0
         
         'ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)
         
         'r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_CtaAux = r_str_CtaAux
         'r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpDeb = g_rst_Princi!BalCom_ImpDeb
         'r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpHab = g_rst_Princi!BalCom_ImpHab
         
         Do While r_int_CadAx1 > 4
               
            ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)
            
            r_int_CadAx2 = r_int_CadAx2 + 2
            r_int_CadAx1 = r_int_CadAx1 - 2
            
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_CtaAux = Mid(r_str_CtaAux, 1, r_int_CadAx1) & String(r_int_CadAx2, "0")
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpDeb = g_rst_Princi!BalCom_ImpDeb
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpHab = g_rst_Princi!BalCom_ImpHab
            
         Loop
         
         If r_int_CadAx1 = 4 Then
         
            ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)
         
            r_int_CadAx2 = r_int_CadAx2 + 1
            r_int_CadAx1 = r_int_CadAx1 - 1
            
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_CtaAux = Mid(r_str_CtaAux, 1, r_int_CadAx1) & String(r_int_CadAx2, "0")
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpDeb = g_rst_Princi!BalCom_ImpDeb
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_ImpHab = g_rst_Princi!BalCom_ImpHab
                        
         End If
         
         g_rst_Princi.MoveNext
         
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   r_Int_NroAux = 1
      
   'Insertando Cuentas a la tabla CTB_BALCOM
   For r_int_NroIte = 1 To UBound(r_arr_CtaAux)
      
      If UBound(r_arr_BalCom) = 0 Then
      
         ReDim Preserve r_arr_BalCom(UBound(r_arr_BalCom) + 1)
         
         r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_CtaAux = r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux
         r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_ImpDeb = r_arr_CtaAux(r_int_NroIte).CtaAux_ImpDeb
         r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_ImpHab = r_arr_CtaAux(r_int_NroIte).CtaAux_ImpHab
         
      Else
      
         Do While r_Int_NroAux <= UBound(r_arr_BalCom)
            
         If CStr(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux)) = CStr(Trim(r_arr_BalCom(r_Int_NroAux).BalCom_CtaAux)) Then
            
            r_arr_BalCom(r_Int_NroAux).BalCom_ImpDeb = r_arr_BalCom(r_Int_NroAux).BalCom_ImpDeb + r_arr_CtaAux(r_int_NroIte).CtaAux_ImpDeb
            r_arr_BalCom(r_Int_NroAux).BalCom_ImpHab = r_arr_BalCom(r_Int_NroAux).BalCom_ImpHab + r_arr_CtaAux(r_int_NroIte).CtaAux_ImpHab
            
            Exit Do
         
         Else
            
            If UBound(r_arr_BalCom) = r_Int_NroAux Then
                              
               ReDim Preserve r_arr_BalCom(UBound(r_arr_BalCom) + 1)
                              
               r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_CtaAux = r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux
               r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_ImpDeb = r_arr_CtaAux(r_int_NroIte).CtaAux_ImpDeb
               r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_ImpHab = r_arr_CtaAux(r_int_NroIte).CtaAux_ImpHab
               
               Exit Do
            
            End If
            
         End If
                        
         r_Int_NroAux = r_Int_NroAux + 1
         
         Loop
                           
      End If
      
      r_Int_NroAux = 1
     
   Next r_int_NroIte
   
   
   'Insertando Cuenta Integradora a la Tabla CTB_BALCOM
   For r_int_NroIte = 1 To UBound(r_arr_BalCom) - 1
           
      r_str_DesCue = ff_DesCue(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux))
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CTB_BALCOM ("
      g_str_Parame = g_str_Parame & "BALCOM_CODEMP, "
      g_str_Parame = g_str_Parame & "BALCOM_PERMES, "
      g_str_Parame = g_str_Parame & "BALCOM_PERANO, "
      g_str_Parame = g_str_Parame & "BALCOM_TIPMON, "
      g_str_Parame = g_str_Parame & "BALCOM_TIPBAL, "
      g_str_Parame = g_str_Parame & "BALCOM_CUENTA, "
      g_str_Parame = g_str_Parame & "BALCOM_DESCRI, "
      g_str_Parame = g_str_Parame & "BALCOM_SLINDB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLINHB, "
      g_str_Parame = g_str_Parame & "BALCOM_IMPDEB, "
      g_str_Parame = g_str_Parame & "BALCOM_IMPHAB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLFIDB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLFIHB, "
      g_str_Parame = g_str_Parame & "BALCOM_FECCRE, "
      g_str_Parame = g_str_Parame & "BALCOM_USUCRE) "
                  
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_PerMes & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_PerAno & "', "
      g_str_Parame = g_str_Parame & 2 & ", "
      g_str_Parame = g_str_Parame & 8 & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_BalCom(r_int_NroIte).BalCom_CtaAux & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_DesCue & "', "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb & ", "
      g_str_Parame = g_str_Parame & r_arr_BalCom(r_int_NroIte).BalCom_ImpHab & ", "
      
      If (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb - r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) > 0 Then
         g_str_Parame = g_str_Parame & (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb - r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) & ", "
      Else
         g_str_Parame = g_str_Parame & 0 & ", "
      End If
      
      If (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb - r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) > 0 Then
         g_str_Parame = g_str_Parame & 0 & ", "
      Else
         If (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb) < (r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) Then
            g_str_Parame = g_str_Parame & (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb - r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) * -1 & ", "
         Else
            g_str_Parame = g_str_Parame & (r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb - r_arr_BalCom(r_int_NroIte).BalCom_ImpHab) & ", "
         End If
      End If
      
      g_str_Parame = g_str_Parame & "'" & l_str_FecImp & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
      
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If
   
   Next r_int_NroIte

End Sub

Private Sub fs_BalCom_MonSol(Optional p_BarPro As SSPanel)
Dim r_dbl_ImpoMN        As Double
Dim r_dbl_ImpoME        As Double
Dim r_str_CodCta        As String
Dim r_str_CtaAux        As String
Dim r_int_FlagDH        As Integer
Dim r_int_ConAux        As Integer
Dim r_int_ConCue        As Integer
Dim r_int_NroIte        As Integer
Dim r_int_Contad        As Integer
Dim r_int_ConVar        As Integer
Dim r_int_Contad2       As Integer
Dim r_dbl_MtoDeb        As Double
Dim r_dbl_MtoHab        As Double
Dim r_int_CadAx1        As Integer
Dim r_int_CadAx2        As Integer
Dim r_int_Iterac        As Integer
Dim r_str_Cuenta        As String
Dim r_str_DesCue        As String
Dim r_Int_NroAux        As Integer
Dim r_dbl_TipCam        As Double
Dim r_dbl_ClDbCu        As Double
Dim r_dbl_ClHbCu        As Double
Dim r_dbl_ClDbCi        As Double
Dim r_dbl_ClHbCi        As Double
Dim r_dbl_MtoEje        As Double
Dim r_arr_CtaCtb()      As modtac_tpo_CtaCtb
Dim r_arr_CtaAux()      As modtac_tpo_CtaAux
Dim r_arr_BalCom()      As modtac_tpo_BalCom

   ReDim r_arr_CtaCtb(0)
   ReDim r_arr_CtaAux(0)
   ReDim r_arr_BalCom(0)

   l_str_FecImp = Format(date, "yyyymmdd")
   l_str_HorImp = Format(Time, "hhmmss")
   p_BarPro.FloodPercent = 0

   'Para leer Cuentas Contables Agrupadas/Obtenemos DEBE-HABER y obtenemos los montos respectivos
   g_str_Parame = "SELECT CNTA_CTBL, FLAG_DEBHAB, SUM(IMP_MOVSOL) AS IMPSOL, SUM(IMP_MOVDOL) AS IMPDOL FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & IIf(l_str_PerMes = 13, l_str_PerMes - 1, l_str_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO <> 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL, FLAG_DEBHAB "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL, FLAG_DEBHAB ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         ReDim Preserve r_arr_CtaCtb(UBound(r_arr_CtaCtb) + 1)
         If IsNull(g_rst_Princi!CNTA_CTBL) Then
            r_str_CodCta = ""
         Else
            r_str_CodCta = Trim(g_rst_Princi!CNTA_CTBL)
         End If

         If g_rst_Princi!FLAG_DEBHAB = "D" Then
            r_int_FlagDH = 1
         ElseIf g_rst_Princi!FLAG_DEBHAB = "H" Then
            r_int_FlagDH = 2
         End If

         If IsNull(g_rst_Princi!IMPSOL) Then
            r_dbl_ImpoMN = 0
         Else
            r_dbl_ImpoMN = g_rst_Princi!IMPSOL
         End If

         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_NumCta = r_str_CodCta
         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_FlagDH = r_int_FlagDH
         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_ImpoMN = r_dbl_ImpoMN
         g_rst_Princi.MoveNext

         l_lng_NumReg = l_lng_NumReg + 1
         DoEvents
         p_BarPro.FloodPercent = CDbl(Format(l_lng_NumReg / l_lng_TotReg * 100, "##0.00"))
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Para leer las Cuentas Contables y obtener el registro de cada cuenta distinta
   g_str_Parame = "SELECT DISTINCT(CNTA_CTBL) FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & IIf(l_str_PerMes = 13, l_str_PerMes - 1, l_str_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO <> 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   r_int_Iterac = 1
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         If Not IsNull(g_rst_Listas!CNTA_CTBL) Then
            r_str_CtaAux = Trim(g_rst_Listas!CNTA_CTBL)
            ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)
            r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_CtaAux = r_str_CtaAux
         End If
         
         g_rst_Listas.MoveNext
         l_lng_NumReg = l_lng_NumReg + 1
         DoEvents
         
         p_BarPro.FloodPercent = CDbl(Format(l_lng_NumReg / l_lng_TotReg * 100, "##0.00"))
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
   
   'Comparacion de Cadena para ingresar el Debe y Haber
   For r_int_ConVar = 1 To UBound(r_arr_CtaCtb)
      For r_int_ConAux = 1 To UBound(r_arr_CtaAux)

         If r_arr_CtaCtb(r_int_ConVar).CtaCtb_NumCta = r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux Then
            If Trim(r_arr_CtaCtb(r_int_ConVar).CtaCtb_FlagDH) = 1 Then
               r_arr_CtaAux(r_int_ConAux).CtaAux_ImpDeb = r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoMN
            Else
               r_arr_CtaAux(r_int_ConAux).CtaAux_ImpHab = r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoMN
            End If
            
            If l_str_PerMes <> 13 Then
            
               If Left(r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux, 1) = 4 Or Left(r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux, 2) = 63 Then
                  If Trim(r_arr_CtaCtb(r_int_ConVar).CtaCtb_FlagDH) = 1 Then
                     r_dbl_ClDbCu = r_dbl_ClDbCu + r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoMN
                  Else
                     r_dbl_ClHbCu = r_dbl_ClHbCu + r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoMN
                  End If
               ElseIf Left(r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux, 1) = 5 Or Left(r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux, 2) = 62 Or Left(r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux, 2) = 64 Or Left(r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux, 2) = 65 Then
                  If Trim(r_arr_CtaCtb(r_int_ConVar).CtaCtb_FlagDH) = 1 Then
                     r_dbl_ClDbCi = r_dbl_ClDbCi + r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoMN
                  Else
                     r_dbl_ClHbCi = r_dbl_ClHbCi + r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoMN
                  End If
               End If
            
            Else
            
               If Left(r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux, 4) = 6911 Or Left(r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux, 4) = 6921 Or Left(r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux, 4) = 6931 Then
                  If Trim(r_arr_CtaCtb(r_int_ConVar).CtaCtb_FlagDH) = 1 Then
                     r_dbl_ClDbCu = r_dbl_ClDbCu + r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoMN
                  Else
                     r_dbl_ClHbCu = r_dbl_ClHbCu + r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoMN
                  End If
               End If
            End If
         End If

      Next r_int_ConAux
   Next r_int_ConVar
   
   If l_str_PerMes <> 13 Then
      
      r_dbl_MtoEje = (r_dbl_ClDbCu - r_dbl_ClHbCu) + (r_dbl_ClDbCi - r_dbl_ClHbCi)
      ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)
      r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux = "391201010101"
      
      If r_dbl_MtoEje >= 0 Then
         r_arr_CtaAux(r_int_ConAux).CtaAux_ImpDeb = r_dbl_MtoEje
      Else
         r_arr_CtaAux(r_int_ConAux).CtaAux_ImpHab = r_dbl_MtoEje
      End If
      
   Else
      
      r_dbl_MtoEje = (r_dbl_ClDbCu - r_dbl_ClHbCu)
      If r_dbl_MtoEje >= 0 Then
         ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)
         r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux = "391102010101"
         r_arr_CtaAux(r_int_ConAux).CtaAux_ImpDeb = r_dbl_MtoEje
      Else
         ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)
         r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux = "391201010101"
         r_arr_CtaAux(r_int_ConAux).CtaAux_ImpHab = r_dbl_MtoEje
      End If
   
   End If

   'Insertando Cuentas a la Tabla CTB_BALCOM
   For r_int_NroIte = 1 To UBound(r_arr_CtaAux)

      r_str_DesCue = ff_DesCue(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux))

      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CTB_BALCOM ("
      g_str_Parame = g_str_Parame & "BALCOM_CODEMP, "
      g_str_Parame = g_str_Parame & "BALCOM_PERMES, "
      g_str_Parame = g_str_Parame & "BALCOM_PERANO, "
      g_str_Parame = g_str_Parame & "BALCOM_TIPMON, "
      g_str_Parame = g_str_Parame & "BALCOM_TIPBAL, "
      g_str_Parame = g_str_Parame & "BALCOM_CUENTA, "
      g_str_Parame = g_str_Parame & "BALCOM_DESCRI, "
      g_str_Parame = g_str_Parame & "BALCOM_SLINDB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLINHB, "
      g_str_Parame = g_str_Parame & "BALCOM_IMPDEB, "
      g_str_Parame = g_str_Parame & "BALCOM_IMPHAB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLFIDB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLFIHB, "
      g_str_Parame = g_str_Parame & "BALCOM_FECCRE, "
      g_str_Parame = g_str_Parame & "BALCOM_USUCRE) "
      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_PerMes & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_PerAno & "', "
      g_str_Parame = g_str_Parame & 1 & ", "
      
      If Left(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux, 1) = 3 Then
         g_str_Parame = g_str_Parame & 9 & ", "
      Else
         If Mid(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux), 3, 1) = 2 Then
            If (Left(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux), 1) = 4 Or Left(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux), 1) = 5) Then
               g_str_Parame = g_str_Parame & 1 & ", "
            Else
               g_str_Parame = g_str_Parame & 3 & ", "
            End If
         Else
            g_str_Parame = g_str_Parame & 1 & ", "
         End If
      End If
      
      g_str_Parame = g_str_Parame & "'" & r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_DesCue & "', "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & r_arr_CtaAux(r_int_NroIte).CtaAux_ImpDeb & ", "
      g_str_Parame = g_str_Parame & r_arr_CtaAux(r_int_NroIte).CtaAux_ImpHab & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_FecImp & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If

   Next r_int_NroIte

End Sub

Private Sub fs_BalCom_MonDol(Optional p_BarPro As SSPanel)

   Dim r_dbl_ImpoMN        As Double
   Dim r_dbl_ImpoME        As Double

   Dim r_str_CodCta        As String
   Dim r_str_CtaAux        As String

   Dim r_int_FlagDH        As Integer
   Dim r_int_ConAux        As Integer
   Dim r_int_ConCue        As Integer
   Dim r_int_NroIte        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConVar        As Integer
   Dim r_int_Contad2       As Integer

   Dim r_dbl_MtoDeb        As Double
   Dim r_dbl_MtoHab        As Double

   Dim r_int_CadAx1        As Integer
   Dim r_int_CadAx2        As Integer
   Dim r_int_Iterac        As Integer

   Dim r_dbl_ActDeb_MN     As Double
   Dim r_dbl_ActDeb_ME     As Double

   Dim r_dbl_ActHab_MN     As Double
   Dim r_dbl_ActHab_ME     As Double

   Dim r_str_Cuenta_MN     As String
   Dim r_str_Cuenta_ME     As String
   Dim r_str_Cuenta        As String
   Dim r_str_DesCue        As String

   Dim r_Int_NroAux        As Integer

   Dim r_dbl_TipCam        As Double

   Dim r_arr_CtaCtb()      As modtac_tpo_CtaCtb
   Dim r_arr_CtaAux()      As modtac_tpo_CtaAux
   Dim r_arr_BalCom()      As modtac_tpo_BalCom

   ReDim r_arr_CtaCtb(0)
   ReDim r_arr_CtaAux(0)
   ReDim r_arr_BalCom(0)

   l_str_FecImp = Format(date, "yyyymmdd")
   l_str_HorImp = Format(Time, "hhmmss")


   'Para leer Cuentas Contables Agrupadas/Obtenemos DEBE-HABER y obtenemos los montos respectivos
   g_str_Parame = "SELECT CNTA_CTBL, FLAG_DEBHAB, SUM(IMP_MOVSOL) AS IMPSOL, SUM(IMP_MOVDOL) AS IMPDOL FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & IIf(l_str_PerMes = 13, l_str_PerMes - 1, l_str_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,3,1)  = 2 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 4 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 5 AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO <> 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL, FLAG_DEBHAB "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL, FLAG_DEBHAB ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF

         ReDim Preserve r_arr_CtaCtb(UBound(r_arr_CtaCtb) + 1)

         r_str_CodCta = Trim(g_rst_Princi!CNTA_CTBL)

         If g_rst_Princi!FLAG_DEBHAB = "D" Then
            r_int_FlagDH = 1
         ElseIf g_rst_Princi!FLAG_DEBHAB = "H" Then
            r_int_FlagDH = 2
         End If

         If IsNull(g_rst_Princi!IMPDOL) Then
            r_dbl_ImpoME = 0
         Else
            r_dbl_ImpoME = g_rst_Princi!IMPDOL
         End If

         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_NumCta = r_str_CodCta
         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_FlagDH = r_int_FlagDH
         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_ImpoME = r_dbl_ImpoME

         g_rst_Princi.MoveNext

         l_lng_NumReg = l_lng_NumReg + 1
         DoEvents

         p_BarPro.FloodPercent = CDbl(Format(l_lng_NumReg / l_lng_TotReg * 100, "##0.00"))

      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing


   'Para leer las Cuentas Contables y obtener el registro de cada cuenta distinta
   g_str_Parame = "SELECT DISTINCT(CNTA_CTBL) FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & IIf(l_str_PerMes = 13, l_str_PerMes - 1, l_str_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,3,1)  = 2 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 4 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 5 AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO <> 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If

   r_int_Iterac = 1

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then

      g_rst_Listas.MoveFirst

      Do While Not g_rst_Listas.EOF

         r_str_CtaAux = Trim(g_rst_Listas!CNTA_CTBL)

         ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)

         r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_CtaAux = r_str_CtaAux

         g_rst_Listas.MoveNext

         l_lng_NumReg = l_lng_NumReg + 1
         DoEvents

         p_BarPro.FloodPercent = CDbl(Format(l_lng_NumReg / l_lng_TotReg * 100, "##0.00"))

      Loop
   End If

   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

   'Comparacion de Cadena para ingresar el Debe y Haber

   For r_int_ConVar = 1 To UBound(r_arr_CtaCtb)

      For r_int_ConAux = 1 To UBound(r_arr_CtaAux)

         If r_arr_CtaCtb(r_int_ConVar).CtaCtb_NumCta = r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux Then

            If Trim(r_arr_CtaCtb(r_int_ConVar).CtaCtb_FlagDH) = 1 Then
               r_arr_CtaAux(r_int_ConAux).CtaAux_ImpDeb = r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoME
            Else
               r_arr_CtaAux(r_int_ConAux).CtaAux_ImpHab = r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoME
            End If

         End If

      Next r_int_ConAux
   Next r_int_ConVar


   'Insertando Cuenta Integradora a la Tabla Temporal TMP_BALCOM
   For r_int_NroIte = 1 To UBound(r_arr_CtaAux)

      r_str_DesCue = ff_DesCue(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux))

      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "INSERT INTO CTB_BALCOM ("
      g_str_Parame = g_str_Parame & "BALCOM_CODEMP, "
      g_str_Parame = g_str_Parame & "BALCOM_PERMES, "
      g_str_Parame = g_str_Parame & "BALCOM_PERANO, "
      g_str_Parame = g_str_Parame & "BALCOM_TIPMON, "
      g_str_Parame = g_str_Parame & "BALCOM_TIPBAL, "
      g_str_Parame = g_str_Parame & "BALCOM_CUENTA, "
      g_str_Parame = g_str_Parame & "BALCOM_DESCRI, "
      g_str_Parame = g_str_Parame & "BALCOM_SLINDB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLINHB, "
      g_str_Parame = g_str_Parame & "BALCOM_IMPDEB, "
      g_str_Parame = g_str_Parame & "BALCOM_IMPHAB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLFIDB, "
      g_str_Parame = g_str_Parame & "BALCOM_SLFIHB, "
      g_str_Parame = g_str_Parame & "BALCOM_FECCRE, "
      g_str_Parame = g_str_Parame & "BALCOM_USUCRE) "

      g_str_Parame = g_str_Parame & "VALUES ("
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_PerMes & "', "
      g_str_Parame = g_str_Parame & "'" & l_str_PerAno & "', "
      g_str_Parame = g_str_Parame & 2 & ", "
      g_str_Parame = g_str_Parame & 2 & ", "
      g_str_Parame = g_str_Parame & "'" & r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_DesCue & "', "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & r_arr_CtaAux(r_int_NroIte).CtaAux_ImpDeb & ", "
      g_str_Parame = g_str_Parame & r_arr_CtaAux(r_int_NroIte).CtaAux_ImpHab & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_FecImp & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If

   Next r_int_NroIte

End Sub

Function ff_DesCue(ByVal p_Descue As String) As String

   g_str_Parame = "SELECT CTAMAE_CODCTA, CTAMAE_DESCRI FROM CTB_CTAMAE WHERE "
   g_str_Parame = g_str_Parame & "CTAMAE_CODCTA = " & p_Descue & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then

      g_rst_Listas.MoveFirst

      Do While Not g_rst_Listas.EOF

         If IsNull(Trim(g_rst_Listas!CTAMAE_DESCRI)) Then
            ff_DesCue = 0
         Else
            ff_DesCue = Trim(g_rst_Listas!CTAMAE_DESCRI)
         End If

         g_rst_Listas.MoveNext

      Loop
   End If

   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

End Function

Private Sub fs_CuePad(ByVal p_TipMon As Integer)

   Dim r_str_DesCue As String

   g_str_Parame = "SELECT DISTINCT(CONCAT(SUBSTR(BALCOM_CUENTA,1,1),'00000000000')) AS CUENTA, SUM(BALCOM_IMPDEB) AS IMPDEB, SUM(BALCOM_IMPHAB) AS IMPHAB FROM CTB_BALCOM WHERE "
   g_str_Parame = g_str_Parame & "SUBSTR(BALCOM_CUENTA,3,10) = '0000000000' AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_TIPMON = " & p_TipMon & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " "
   g_str_Parame = g_str_Parame & "GROUP BY CONCAT(SUBSTR(BALCOM_CUENTA,1,1),'00000000000') "
   g_str_Parame = g_str_Parame & "ORDER BY CONCAT(SUBSTR(BALCOM_CUENTA,1,1),'00000000000') ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF

         r_str_DesCue = ff_DesCue(Trim(g_rst_Princi!Cuenta))

         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO CTB_BALCOM ("
         g_str_Parame = g_str_Parame & "BALCOM_CODEMP, "
         g_str_Parame = g_str_Parame & "BALCOM_PERMES, "
         g_str_Parame = g_str_Parame & "BALCOM_PERANO, "
         g_str_Parame = g_str_Parame & "BALCOM_TIPMON, "
         g_str_Parame = g_str_Parame & "BALCOM_TIPBAL, "
         g_str_Parame = g_str_Parame & "BALCOM_CUENTA, "
         g_str_Parame = g_str_Parame & "BALCOM_DESCRI, "
         g_str_Parame = g_str_Parame & "BALCOM_SLINDB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLINHB, "
         g_str_Parame = g_str_Parame & "BALCOM_IMPDEB, "
         g_str_Parame = g_str_Parame & "BALCOM_IMPHAB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLFIDB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLFIHB, "
         g_str_Parame = g_str_Parame & "BALCOM_FECCRE, "
         g_str_Parame = g_str_Parame & "BALCOM_USUCRE) "

         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_PerMes & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_PerAno & "', "
         g_str_Parame = g_str_Parame & p_TipMon & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!Cuenta) & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesCue & "', "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!IMPDEB & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!IMPHAB & ", "
         
         If (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) > 0 Then
            g_str_Parame = g_str_Parame & (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) & ", "
         Else
            g_str_Parame = g_str_Parame & 0 & ", "
         End If
         
         If (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) > 0 Then
            g_str_Parame = g_str_Parame & 0 & ", "
         Else
            If (g_rst_Princi!IMPDEB) < (g_rst_Princi!IMPHAB) Then
               g_str_Parame = g_str_Parame & (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) * -1 & ", "
            Else
               g_str_Parame = g_str_Parame & (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) & ", "
            End If
         End If
         
         g_str_Parame = g_str_Parame & "'" & l_str_FecImp & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "

         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If

         g_rst_Princi.MoveNext

      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Sub

Private Sub fs_CueInt(ByVal p_TipMon As Integer)

   Dim r_str_DesCue  As String

   g_str_Parame = "SELECT CONCAT(CONCAT(SUBSTR(BALCOM_CUENTA,1,2),'0'),SUBSTR(BALCOM_CUENTA,4,12)) AS CUENTA, SUM(BALCOM_IMPDEB) AS IMPDEB, SUM(BALCOM_IMPHAB) AS IMPHAB FROM CTB_BALCOM WHERE "
   g_str_Parame = g_str_Parame & "BALCOM_TIPMON = " & p_TipMon & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " "
   g_str_Parame = g_str_Parame & "GROUP BY CONCAT(CONCAT(SUBSTR(BALCOM_CUENTA,1,2),'0'),SUBSTR(BALCOM_CUENTA,4,12)) "
   g_str_Parame = g_str_Parame & "ORDER BY CONCAT(CONCAT(SUBSTR(BALCOM_CUENTA,1,2),'0'),SUBSTR(BALCOM_CUENTA,4,12)) ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF

         r_str_DesCue = ff_DesCue(Trim(g_rst_Princi!Cuenta))

         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO CTB_BALCOM ("
         g_str_Parame = g_str_Parame & "BALCOM_CODEMP, "
         g_str_Parame = g_str_Parame & "BALCOM_PERMES, "
         g_str_Parame = g_str_Parame & "BALCOM_PERANO, "
         g_str_Parame = g_str_Parame & "BALCOM_TIPMON, "
         g_str_Parame = g_str_Parame & "BALCOM_TIPBAL, "
         g_str_Parame = g_str_Parame & "BALCOM_CUENTA, "
         g_str_Parame = g_str_Parame & "BALCOM_DESCRI, "
         g_str_Parame = g_str_Parame & "BALCOM_SLINDB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLINHB, "
         g_str_Parame = g_str_Parame & "BALCOM_IMPDEB, "
         g_str_Parame = g_str_Parame & "BALCOM_IMPHAB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLFIDB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLFIHB, "
         g_str_Parame = g_str_Parame & "BALCOM_FECCRE, "
         g_str_Parame = g_str_Parame & "BALCOM_USUCRE) "

         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_PerMes & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_PerAno & "', "
         g_str_Parame = g_str_Parame & p_TipMon & ", "
         
         If Left(Trim(g_rst_Princi!Cuenta), 1) = 3 Then
            g_str_Parame = g_str_Parame & 9 & ", "
         Else
            g_str_Parame = g_str_Parame & 0 & ", "
         End If
         
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!Cuenta) & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesCue & "', "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!IMPDEB & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!IMPHAB & ", "
         
         If (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) > 0 Then
            g_str_Parame = g_str_Parame & (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) & ", "
         Else
            g_str_Parame = g_str_Parame & 0 & ", "
         End If
         
         If (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) > 0 Then
            g_str_Parame = g_str_Parame & 0 & ", "
         Else
            If (g_rst_Princi!IMPDEB) < (g_rst_Princi!IMPHAB) Then
               g_str_Parame = g_str_Parame & (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) * -1 & ", "
            Else
               g_str_Parame = g_str_Parame & (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) & ", "
            End If
         End If
      
         g_str_Parame = g_str_Parame & "'" & l_str_FecImp & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "

         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If

         g_rst_Princi.MoveNext

      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Sub

Private Sub fs_CueFal(ByVal p_TipMon As Integer)

   Dim r_str_DesCue        As String

   g_str_Parame = "SELECT CONCAT(CONCAT(SUBSTR(BALCOM_CUENTA,1,1),'0'),SUBSTR(BALCOM_CUENTA,3,12)) AS CUENTA, SUM(BALCOM_IMPDEB) AS IMPDEB, SUM(BALCOM_IMPHAB) AS IMPHAB FROM CTB_BALCOM WHERE "
   g_str_Parame = g_str_Parame & "BALCOM_TIPMON = " & p_TipMon & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " AND "
   g_str_Parame = g_str_Parame & "SUBSTR(BALCOM_CUENTA,4,9) = '000000000' AND "
   g_str_Parame = g_str_Parame & "SUBSTR(BALCOM_CUENTA,1,3) < 200 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(BALCOM_CUENTA,3,1) = 1 "
   g_str_Parame = g_str_Parame & "GROUP BY CONCAT(CONCAT(SUBSTR(BALCOM_CUENTA,1,1),'0'),SUBSTR(BALCOM_CUENTA,3,12)) "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF

         r_str_DesCue = ff_DesCue(Trim(g_rst_Princi!Cuenta))

         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO CTB_BALCOM ("
         g_str_Parame = g_str_Parame & "BALCOM_CODEMP, "
         g_str_Parame = g_str_Parame & "BALCOM_PERMES, "
         g_str_Parame = g_str_Parame & "BALCOM_PERANO, "
         g_str_Parame = g_str_Parame & "BALCOM_TIPMON, "
         g_str_Parame = g_str_Parame & "BALCOM_TIPBAL, "
         g_str_Parame = g_str_Parame & "BALCOM_CUENTA, "
         g_str_Parame = g_str_Parame & "BALCOM_DESCRI, "
         g_str_Parame = g_str_Parame & "BALCOM_SLINDB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLINHB, "
         g_str_Parame = g_str_Parame & "BALCOM_IMPDEB, "
         g_str_Parame = g_str_Parame & "BALCOM_IMPHAB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLFIDB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLFIHB, "
         g_str_Parame = g_str_Parame & "BALCOM_FECCRE, "
         g_str_Parame = g_str_Parame & "BALCOM_USUCRE) "

         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_PerMes & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_PerAno & "', "
         g_str_Parame = g_str_Parame & p_TipMon & ", "
         g_str_Parame = g_str_Parame & 0 & ","
         g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!Cuenta) & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesCue & "', "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!IMPDEB & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!IMPHAB & ", "
         
         If (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) > 0 Then
            g_str_Parame = g_str_Parame & (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) & ", "
         Else
            g_str_Parame = g_str_Parame & 0 & ", "
         End If
         
         If (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) > 0 Then
            g_str_Parame = g_str_Parame & 0 & ", "
         Else
            If (g_rst_Princi!IMPDEB) < (g_rst_Princi!IMPHAB) Then
               g_str_Parame = g_str_Parame & (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) * -1 & ", "
            Else
               g_str_Parame = g_str_Parame & (g_rst_Princi!IMPDEB - g_rst_Princi!IMPHAB) & ", "
            End If
         End If
         
         g_str_Parame = g_str_Parame & "'" & l_str_FecImp & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "

         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If

         g_rst_Princi.MoveNext

      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Sub

Private Sub fs_SalEne_MonSol(ByVal p_TipMon As Integer)

   Dim r_dbl_ImpoMN        As Double
   Dim r_dbl_ImpoME        As Double

   Dim r_str_CodCta        As String
   Dim r_str_CtaAux        As String

   Dim r_int_FlagDH        As Integer
   Dim r_int_ConAux        As Integer
   Dim r_int_ConCue        As Integer
   Dim r_int_NroIte        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConVar        As Integer
   Dim r_int_Contad2       As Integer

   Dim r_dbl_MtoDeb        As Double
   Dim r_dbl_MtoHab        As Double

   Dim r_int_CadAx1        As Integer
   Dim r_int_CadAx2        As Integer
   Dim r_int_Iterac        As Integer

   Dim r_str_Cuenta        As String
   Dim r_str_DesCue        As String

   Dim r_Int_NroAux        As Integer
   Dim r_int_NumIte        As Integer

   Dim r_dbl_TipCam        As Double
   Dim r_int_ConFlg        As Integer

   Dim r_arr_CtaCtb()      As modtac_tpo_CtaCtb
   Dim r_arr_CtaAux()      As modtac_tpo_CtaAux
   Dim r_arr_BalCom()      As modtac_tpo_BalCom

   ReDim r_arr_CtaCtb(0)
   ReDim r_arr_CtaAux(0)
   ReDim r_arr_BalCom(0)

   l_str_FecImp = Format(date, "yyyymmdd")
   l_str_HorImp = Format(Time, "hhmmss")
    
   'Llenado del libro 5 al Arreglo para la Apertura (Solo Enero)
   g_str_Parame = "SELECT CNTA_CTBL, FLAG_DEBHAB, SUM(IMP_MOVSOL) AS IMPSOL, SUM(IMP_MOVDOL) AS IMPDOL FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO = 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL, FLAG_DEBHAB "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL, FLAG_DEBHAB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

      g_rst_Princi.MoveFirst
      
      ReDim r_arr_CtaCtb(0)

      Do While Not g_rst_Princi.EOF

         ReDim Preserve r_arr_CtaCtb(UBound(r_arr_CtaCtb) + 1)
         
         r_str_CodCta = Trim(g_rst_Princi!CNTA_CTBL)

         If g_rst_Princi!FLAG_DEBHAB = "D" Then
            r_int_FlagDH = 1
         ElseIf g_rst_Princi!FLAG_DEBHAB = "H" Then
            r_int_FlagDH = 2
         End If
                           
         If IsNull(g_rst_Princi!IMPSOL) Then
            r_dbl_ImpoMN = 0
         Else
            r_dbl_ImpoMN = g_rst_Princi!IMPSOL
         End If

         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_NumCta = r_str_CodCta
         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_FlagDH = r_int_FlagDH
         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_ImpoMN = r_dbl_ImpoMN
      
         g_rst_Princi.MoveNext

         l_lng_NumReg = l_lng_NumReg + 1
         DoEvents

         'p_BarPro.FloodPercent = CDbl(Format(l_lng_NumReg / l_lng_TotReg * 100, "##0.00"))
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing


   'Para leer las Cuentas Contables y obtener el registro de cada cuenta distinta
   g_str_Parame = "SELECT DISTINCT(CNTA_CTBL) FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO = 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If

   r_int_Iterac = 1

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then

      g_rst_Listas.MoveFirst
      
      ReDim r_arr_CtaAux(0)

      Do While Not g_rst_Listas.EOF

         r_str_CtaAux = Trim(g_rst_Listas!CNTA_CTBL)

         ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)

         r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_CtaAux = r_str_CtaAux

         g_rst_Listas.MoveNext

         l_lng_NumReg = l_lng_NumReg + 1
         DoEvents

         'p_BarPro.FloodPercent = CDbl(Format(l_lng_NumReg / l_lng_TotReg * 100, "##0.00"))

      Loop
   End If


   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

   'Comparacion de Cadena para ingresar el Debe y Haber

   For r_int_ConVar = 1 To UBound(r_arr_CtaCtb)

      For r_int_ConAux = 1 To UBound(r_arr_CtaAux)

         If r_arr_CtaCtb(r_int_ConVar).CtaCtb_NumCta = r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux Then

            If Trim(r_arr_CtaCtb(r_int_ConVar).CtaCtb_FlagDH) = 1 Then
               r_arr_CtaAux(r_int_ConAux).CtaAux_ImpDeb = r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoMN
            Else
               r_arr_CtaAux(r_int_ConAux).CtaAux_ImpHab = r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoMN
            End If

         End If

      Next r_int_ConAux
   Next r_int_ConVar
   
   
   'Llenado del Arreglo a la Tabla CTB_BALCOM
   For r_int_NroIte = 1 To UBound(r_arr_CtaAux)

      g_str_Parame = "SELECT * FROM CTB_BALCOM WHERE "
      g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_TIPMON = " & p_TipMon & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_CUENTA = '" & Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux) & "' "
      g_str_Parame = g_str_Parame & "ORDER BY BALCOM_TIPMON ASC, BALCOM_CUENTA ASC "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      'UPDATE
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
                   
         For r_int_NumIte = 1 To UBound(r_arr_CtaAux)
         
            r_int_ConFlg = 0
            
            For r_Int_NroAux = 1 To UBound(r_arr_CtaAux)

               If r_arr_CtaAux(r_int_NumIte).CtaAux_CtaAux = Trim(g_rst_Princi!BALCOM_CUENTA) Then
                  
                  r_int_ConFlg = 1
                  
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & "UPDATE CTB_BALCOM SET "
                  g_str_Parame = g_str_Parame & "BALCOM_SLINDB = " & r_arr_CtaAux(r_int_NumIte).CtaAux_ImpDeb & ", "
                  g_str_Parame = g_str_Parame & "BALCOM_SLINHB = " & r_arr_CtaAux(r_int_NumIte).CtaAux_ImpHab & " "
                  g_str_Parame = g_str_Parame & "WHERE "
                  g_str_Parame = g_str_Parame & "BALCOM_CODEMP = '" & l_str_CodEmp & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_FECCRE = '" & l_str_FecImp & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_TIPMON = '" & p_TipMon & "' AND "
                  
                  If Mid(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux), 3, 1) = 2 Then
                     If (Left(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux), 1) = 4 Or Left(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux), 1) = 5) Then
                        g_str_Parame = g_str_Parame & "BALCOM_TIPBAL = '" & 1 & "' AND "
                     Else
                        g_str_Parame = g_str_Parame & "BALCOM_TIPBAL = '" & 3 & "' AND "
                     End If
                  Else
                     g_str_Parame = g_str_Parame & "BALCOM_TIPBAL = '" & 1 & "' AND "
                  End If
                  
                  g_str_Parame = g_str_Parame & "BALCOM_USUCRE = '" & modgen_g_str_CodUsu & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_PERMES = '" & l_str_PerMes & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_PERANO = '" & l_str_PerAno & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_CUENTA = '" & Trim(g_rst_Princi!BALCOM_CUENTA) & "' "

                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                     Exit Sub
                  End If
                  
                  Exit For
                  
               End If
                  
            Next r_Int_NroAux
         Next r_int_NumIte
      'INSERT
      Else
         
         g_str_Parame = "SELECT DISTINCT(CNTA_CTBL) AS CUENTA FROM CNTBL_ASIENTO_DET WHERE "
         g_str_Parame = g_str_Parame & "MES = " & l_str_PerMes & " AND "
         g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
         g_str_Parame = g_str_Parame & "NRO_LIBRO = 5 "
         g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL ASC "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If

            If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

               g_rst_Princi.MoveFirst

               Do While Not g_rst_Princi.EOF

                  If Trim(g_rst_Princi!Cuenta) = Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux) Then
                     
                     r_str_DesCue = ff_DesCue(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux))

                     g_str_Parame = ""
                     g_str_Parame = g_str_Parame & "INSERT INTO CTB_BALCOM ("
                     g_str_Parame = g_str_Parame & "BALCOM_CODEMP, "
                     g_str_Parame = g_str_Parame & "BALCOM_PERMES, "
                     g_str_Parame = g_str_Parame & "BALCOM_PERANO, "
                     g_str_Parame = g_str_Parame & "BALCOM_TIPMON, "
                     g_str_Parame = g_str_Parame & "BALCOM_TIPBAL, "
                     g_str_Parame = g_str_Parame & "BALCOM_CUENTA, "
                     g_str_Parame = g_str_Parame & "BALCOM_DESCRI, "
                     g_str_Parame = g_str_Parame & "BALCOM_SLINDB, "
                     g_str_Parame = g_str_Parame & "BALCOM_SLINHB, "
                     g_str_Parame = g_str_Parame & "BALCOM_IMPDEB, "
                     g_str_Parame = g_str_Parame & "BALCOM_IMPHAB, "
                     g_str_Parame = g_str_Parame & "BALCOM_SLFIDB, "
                     g_str_Parame = g_str_Parame & "BALCOM_SLFIHB, "
                     g_str_Parame = g_str_Parame & "BALCOM_FECCRE, "
                     g_str_Parame = g_str_Parame & "BALCOM_USUCRE) "

                     g_str_Parame = g_str_Parame & "VALUES ("
                     g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
                     g_str_Parame = g_str_Parame & "'" & l_str_PerMes & "', "
                     g_str_Parame = g_str_Parame & "'" & l_str_PerAno & "', "
                     g_str_Parame = g_str_Parame & p_TipMon & ", "

                     If Mid(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux), 3, 1) = 2 Then
                        If (Left(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux), 1) = 4 Or Left(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux), 1) = 5) Then
                           g_str_Parame = g_str_Parame & 1 & ", "
                        Else
                           g_str_Parame = g_str_Parame & 3 & ", "
                        End If
                     Else
                        g_str_Parame = g_str_Parame & 1 & ", "
                     End If
                     
                     g_str_Parame = g_str_Parame & "'" & Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux) & "', "
                     g_str_Parame = g_str_Parame & "'" & r_str_DesCue & "', "
                     g_str_Parame = g_str_Parame & r_arr_CtaAux(r_int_NroIte).CtaAux_ImpDeb & ", "
                     g_str_Parame = g_str_Parame & r_arr_CtaAux(r_int_NroIte).CtaAux_ImpHab & ", "
                     g_str_Parame = g_str_Parame & 0 & ", "
                     g_str_Parame = g_str_Parame & 0 & ", "
                     g_str_Parame = g_str_Parame & 0 & ", "
                     g_str_Parame = g_str_Parame & 0 & ", "
                     g_str_Parame = g_str_Parame & "'" & l_str_FecImp & "', "
                     g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "

                     If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                        Exit Sub
                     End If
                  
                  End If

                  g_rst_Princi.MoveNext

               Loop

            End If

            g_rst_Princi.Close
            Set g_rst_Princi = Nothing
  
      End If
      
   Next r_int_NroIte


End Sub

Private Sub fs_SalEne_MonDol(ByVal p_TipMon As Integer)

   Dim r_dbl_ImpoMN        As Double
   Dim r_dbl_ImpoME        As Double

   Dim r_str_CodCta        As String
   Dim r_str_CtaAux        As String

   Dim r_int_FlagDH        As Integer
   Dim r_int_ConAux        As Integer
   Dim r_int_ConCue        As Integer
   Dim r_int_NroIte        As Integer
   Dim r_int_Contad        As Integer
   Dim r_int_ConVar        As Integer
   Dim r_int_Contad2       As Integer

   Dim r_dbl_MtoDeb        As Double
   Dim r_dbl_MtoHab        As Double

   Dim r_int_CadAx1        As Integer
   Dim r_int_CadAx2        As Integer
   Dim r_int_Iterac        As Integer

   Dim r_str_Cuenta        As String
   Dim r_str_DesCue        As String

   Dim r_Int_NroAux        As Integer
   Dim r_int_NumIte        As Integer

   Dim r_dbl_TipCam        As Double
   Dim r_int_ConFlg        As Integer

   Dim r_arr_CtaCtb()      As modtac_tpo_CtaCtb
   Dim r_arr_CtaAux()      As modtac_tpo_CtaAux
   Dim r_arr_BalCom()      As modtac_tpo_BalCom

   ReDim r_arr_CtaCtb(0)
   ReDim r_arr_CtaAux(0)
   ReDim r_arr_BalCom(0)

   l_str_FecImp = Format(date, "yyyymmdd")
   l_str_HorImp = Format(Time, "hhmmss")
    
   'Llenado del libro 5 al Arreglo para la Apertura (Solo Enero)
   g_str_Parame = "SELECT CNTA_CTBL, FLAG_DEBHAB, SUM(IMP_MOVSOL) AS IMPSOL, SUM(IMP_MOVDOL) AS IMPDOL FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,3,1)  = 2 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 4 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 5 AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO = 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL, FLAG_DEBHAB "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL, FLAG_DEBHAB ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

      g_rst_Princi.MoveFirst
      
      ReDim r_arr_CtaCtb(0)

      Do While Not g_rst_Princi.EOF

         ReDim Preserve r_arr_CtaCtb(UBound(r_arr_CtaCtb) + 1)
         
         r_str_CodCta = Trim(g_rst_Princi!CNTA_CTBL)

         If g_rst_Princi!FLAG_DEBHAB = "D" Then
            r_int_FlagDH = 1
         ElseIf g_rst_Princi!FLAG_DEBHAB = "H" Then
            r_int_FlagDH = 2
         End If
                           
         If IsNull(g_rst_Princi!IMPDOL) Then
            r_dbl_ImpoME = 0
         Else
            r_dbl_ImpoME = g_rst_Princi!IMPDOL
         End If

         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_NumCta = r_str_CodCta
         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_FlagDH = r_int_FlagDH
         r_arr_CtaCtb(UBound(r_arr_CtaCtb)).CtaCtb_ImpoME = r_dbl_ImpoME
      
         g_rst_Princi.MoveNext

         l_lng_NumReg = l_lng_NumReg + 1
         DoEvents

         'p_BarPro.FloodPercent = CDbl(Format(l_lng_NumReg / l_lng_TotReg * 100, "##0.00"))
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing


   'Para leer las Cuentas Contables y obtener el registro de cada cuenta distinta
   g_str_Parame = "SELECT DISTINCT(CNTA_CTBL) FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,3,1)  = 2 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 4 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 5 AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO = 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If

   r_int_Iterac = 1

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then

      g_rst_Listas.MoveFirst
      
      ReDim r_arr_CtaAux(0)

      Do While Not g_rst_Listas.EOF

         r_str_CtaAux = Trim(g_rst_Listas!CNTA_CTBL)

         ReDim Preserve r_arr_CtaAux(UBound(r_arr_CtaAux) + 1)

         r_arr_CtaAux(UBound(r_arr_CtaAux)).CtaAux_CtaAux = r_str_CtaAux

         g_rst_Listas.MoveNext

         l_lng_NumReg = l_lng_NumReg + 1
         DoEvents

         'p_BarPro.FloodPercent = CDbl(Format(l_lng_NumReg / l_lng_TotReg * 100, "##0.00"))

      Loop
   End If


   g_rst_Listas.Close
   Set g_rst_Listas = Nothing

   'Comparacion de Cadena para ingresar el Debe y Haber

   For r_int_ConVar = 1 To UBound(r_arr_CtaCtb)

      For r_int_ConAux = 1 To UBound(r_arr_CtaAux)

         If r_arr_CtaCtb(r_int_ConVar).CtaCtb_NumCta = r_arr_CtaAux(r_int_ConAux).CtaAux_CtaAux Then

            If Trim(r_arr_CtaCtb(r_int_ConVar).CtaCtb_FlagDH) = 1 Then
               r_arr_CtaAux(r_int_ConAux).CtaAux_ImpDeb = r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoME
            Else
               r_arr_CtaAux(r_int_ConAux).CtaAux_ImpHab = r_arr_CtaCtb(r_int_ConVar).CtaCtb_ImpoME
            End If

         End If

      Next r_int_ConAux
   Next r_int_ConVar
   
   
   'Llenado del Arreglo a la Tabla CTB_BALCOM
   For r_int_NroIte = 1 To UBound(r_arr_CtaAux)

      g_str_Parame = "SELECT * FROM CTB_BALCOM WHERE "
      g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_TIPMON = " & p_TipMon & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_CUENTA = '" & Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux) & "' "
      g_str_Parame = g_str_Parame & "ORDER BY BALCOM_TIPMON ASC, BALCOM_CUENTA ASC "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      'UPDATE
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
                   
         For r_int_NumIte = 1 To UBound(r_arr_CtaAux)
         
            r_int_ConFlg = 0
            
            For r_Int_NroAux = 1 To UBound(r_arr_CtaAux)

               If r_arr_CtaAux(r_int_NumIte).CtaAux_CtaAux = Trim(g_rst_Princi!BALCOM_CUENTA) Then
                  
                  r_int_ConFlg = 1
                  
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & "UPDATE CTB_BALCOM SET "
                  g_str_Parame = g_str_Parame & "BALCOM_SLINDB = " & r_arr_CtaAux(r_int_NumIte).CtaAux_ImpDeb & ", "
                  g_str_Parame = g_str_Parame & "BALCOM_SLINHB = " & r_arr_CtaAux(r_int_NumIte).CtaAux_ImpHab & " "
                  g_str_Parame = g_str_Parame & "WHERE "
                  g_str_Parame = g_str_Parame & "BALCOM_CODEMP = '" & l_str_CodEmp & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_FECCRE = '" & l_str_FecImp & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_TIPMON = '" & p_TipMon & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_TIPBAL = '" & 2 & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_USUCRE = '" & modgen_g_str_CodUsu & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_PERMES = '" & l_str_PerMes & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_PERANO = '" & l_str_PerAno & "' AND "
                  g_str_Parame = g_str_Parame & "BALCOM_CUENTA = '" & Trim(g_rst_Princi!BALCOM_CUENTA) & "' "

                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                     Exit Sub
                  End If
                  
                  Exit For
                  
               End If
                  
            Next r_Int_NroAux
         Next r_int_NumIte
      'INSERT
      Else
         
         g_str_Parame = "SELECT DISTINCT(CNTA_CTBL) AS CUENTA FROM CNTBL_ASIENTO_DET WHERE "
         g_str_Parame = g_str_Parame & "MES = " & l_str_PerMes & " AND "
         g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
         g_str_Parame = g_str_Parame & "NRO_LIBRO = 5 "
         g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL ASC "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If

            If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

               g_rst_Princi.MoveFirst

               Do While Not g_rst_Princi.EOF

                  If Trim(g_rst_Princi!Cuenta) = Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux) Then
                     
                     r_str_DesCue = ff_DesCue(Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux))

                     g_str_Parame = ""
                     g_str_Parame = g_str_Parame & "INSERT INTO CTB_BALCOM ("
                     g_str_Parame = g_str_Parame & "BALCOM_CODEMP, "
                     g_str_Parame = g_str_Parame & "BALCOM_PERMES, "
                     g_str_Parame = g_str_Parame & "BALCOM_PERANO, "
                     g_str_Parame = g_str_Parame & "BALCOM_TIPMON, "
                     g_str_Parame = g_str_Parame & "BALCOM_TIPBAL, "
                     g_str_Parame = g_str_Parame & "BALCOM_CUENTA, "
                     g_str_Parame = g_str_Parame & "BALCOM_DESCRI, "
                     g_str_Parame = g_str_Parame & "BALCOM_SLINDB, "
                     g_str_Parame = g_str_Parame & "BALCOM_SLINHB, "
                     g_str_Parame = g_str_Parame & "BALCOM_IMPDEB, "
                     g_str_Parame = g_str_Parame & "BALCOM_IMPHAB, "
                     g_str_Parame = g_str_Parame & "BALCOM_SLFIDB, "
                     g_str_Parame = g_str_Parame & "BALCOM_SLFIHB, "
                     g_str_Parame = g_str_Parame & "BALCOM_FECCRE, "
                     g_str_Parame = g_str_Parame & "BALCOM_USUCRE) "

                     g_str_Parame = g_str_Parame & "VALUES ("
                     g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
                     g_str_Parame = g_str_Parame & "'" & l_str_PerMes & "', "
                     g_str_Parame = g_str_Parame & "'" & l_str_PerAno & "', "
                     g_str_Parame = g_str_Parame & p_TipMon & ", "
                     g_str_Parame = g_str_Parame & 2 & ", "
                     g_str_Parame = g_str_Parame & "'" & Trim(r_arr_CtaAux(r_int_NroIte).CtaAux_CtaAux) & "', "
                     g_str_Parame = g_str_Parame & "'" & r_str_DesCue & "', "
                     g_str_Parame = g_str_Parame & r_arr_CtaAux(r_int_NroIte).CtaAux_ImpDeb & ", "
                     g_str_Parame = g_str_Parame & r_arr_CtaAux(r_int_NroIte).CtaAux_ImpHab & ", "
                     g_str_Parame = g_str_Parame & 0 & ", "
                     g_str_Parame = g_str_Parame & 0 & ", "
                     g_str_Parame = g_str_Parame & 0 & ", "
                     g_str_Parame = g_str_Parame & 0 & ", "
                     g_str_Parame = g_str_Parame & "'" & l_str_FecImp & "', "
                     g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "

                     If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                        Exit Sub
                     End If
                  
                  End If

                  g_rst_Princi.MoveNext

               Loop

            End If

            g_rst_Princi.Close
            Set g_rst_Princi = Nothing
  
      End If
      
   Next r_int_NroIte


End Sub


Private Sub fs_SalAnt(ByVal p_TipMon As Integer)

   Dim r_int_NroIte        As Integer
  ' Dim r_int_VarAux        As Integer
  ' Dim r_int_VarTem        As Integer
   Dim r_str_CodCta        As String
   Dim r_dbl_SalAct        As Double
   Dim r_int_FlagDH        As Integer
   Dim r_dbl_ImpoMN        As Double
   Dim r_str_CtaAux        As String

   Dim r_str_DesCue        As String
   Dim r_int_Iterac        As Integer
   Dim r_int_ConVar        As Integer
   Dim r_int_ConAux        As Integer
   Dim r_int_NumIte        As Integer
   Dim r_int_ConFlg        As Integer
   Dim r_Int_NroAux        As Integer

   Dim r_arr_BalCom()      As modtac_tpo_BalCom
   
   Dim r_arr_CtaCtb()      As modtac_tpo_CtaCtb
   Dim r_arr_CtaAux()      As modtac_tpo_CtaAux

   ReDim r_arr_BalCom(0)


   'Llenado de la Data de la TABLA CTB_BALCOM al Arreglo
   g_str_Parame = "SELECT * FROM CTB_BALCOM WHERE "

   If CInt(l_str_PerMes) - 1 = 0 Then
      g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & 12 & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & CStr(CInt(l_str_PerAno) - 1) & " AND "
   Else
      If l_str_PerMes = 13 Then
         g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & CStr(CInt(l_str_PerMes) - 2) & " AND "
         g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
      Else
         g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & CStr(CInt(l_str_PerMes) - 1) & " AND "
         g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
      End If
   End If
   g_str_Parame = g_str_Parame & "SUBSTR(TRIM(BALCOM_CUENTA),LENGTH(TRIM(BALCOM_CUENTA))-1,2) <> 00 AND "
   g_str_Parame = g_str_Parame & "BALCOM_TIPBAL <> 9 AND "
   g_str_Parame = g_str_Parame & "BALCOM_TIPMON = " & p_TipMon & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " "
   g_str_Parame = g_str_Parame & "ORDER BY BALCOM_TIPMON ASC, BALCOM_CUENTA ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF

      ReDim Preserve r_arr_BalCom(UBound(r_arr_BalCom) + 1)

      r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_CtaAux = Trim(g_rst_Princi!BALCOM_CUENTA)
      r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_ImpDeb = g_rst_Princi!BALCOM_SLFIDB
      r_arr_BalCom(UBound(r_arr_BalCom)).BalCom_ImpHab = g_rst_Princi!BALCOM_SLFIHB

      g_rst_Princi.MoveNext

      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   

   'Llenado del Arreglo a la Tabla CTB_BALCOM
   For r_int_NroIte = 1 To UBound(r_arr_BalCom)

      g_str_Parame = "SELECT * FROM CTB_BALCOM WHERE "
      g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_TIPMON = " & p_TipMon & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " AND "
      g_str_Parame = g_str_Parame & "BALCOM_CUENTA = '" & Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux) & "' "
      g_str_Parame = g_str_Parame & "ORDER BY BALCOM_TIPMON ASC, BALCOM_CUENTA ASC "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If

      'UPDATE
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "UPDATE CTB_BALCOM SET "
         g_str_Parame = g_str_Parame & "BALCOM_SLINDB = " & r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb & ", "
         g_str_Parame = g_str_Parame & "BALCOM_SLINHB = " & r_arr_BalCom(r_int_NroIte).BalCom_ImpHab & " "
         g_str_Parame = g_str_Parame & "WHERE "
         g_str_Parame = g_str_Parame & "BALCOM_CODEMP = '" & l_str_CodEmp & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_FECCRE = '" & l_str_FecImp & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_TIPMON = '" & p_TipMon & "' AND "

         If Mid(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 3, 1) <> 0 Then
            If p_TipMon = 1 Then
               If Mid(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 3, 1) = 2 Then
                  If (Left(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 1) = 4 Or Left(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 1) = 5) Then
                     g_str_Parame = g_str_Parame & "BALCOM_TIPBAL = '" & 1 & "' AND "
                  Else
                     g_str_Parame = g_str_Parame & "BALCOM_TIPBAL = '" & 3 & "' AND "
                  End If
               Else
                  g_str_Parame = g_str_Parame & "BALCOM_TIPBAL = '" & 1 & "' AND "
               End If
            ElseIf p_TipMon = 2 Then
               g_str_Parame = g_str_Parame & "BALCOM_TIPBAL = '" & 2 & "' AND "
             End If
         Else
            g_str_Parame = g_str_Parame & "BALCOM_TIPBAL = '" & 0 & "' AND "
         End If

         g_str_Parame = g_str_Parame & "BALCOM_USUCRE = '" & modgen_g_str_CodUsu & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_PERMES = '" & l_str_PerMes & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_PERANO = '" & l_str_PerAno & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_CUENTA = '" & Trim(g_rst_Princi!BALCOM_CUENTA) & "' "

         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
 
     'INSERT
      Else
         
         r_str_DesCue = ff_DesCue(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux))

         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO CTB_BALCOM ("
         g_str_Parame = g_str_Parame & "BALCOM_CODEMP, "
         g_str_Parame = g_str_Parame & "BALCOM_PERMES, "
         g_str_Parame = g_str_Parame & "BALCOM_PERANO, "
         g_str_Parame = g_str_Parame & "BALCOM_TIPMON, "
         g_str_Parame = g_str_Parame & "BALCOM_TIPBAL, "
         g_str_Parame = g_str_Parame & "BALCOM_CUENTA, "
         g_str_Parame = g_str_Parame & "BALCOM_DESCRI, "
         g_str_Parame = g_str_Parame & "BALCOM_SLINDB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLINHB, "
         g_str_Parame = g_str_Parame & "BALCOM_IMPDEB, "
         g_str_Parame = g_str_Parame & "BALCOM_IMPHAB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLFIDB, "
         g_str_Parame = g_str_Parame & "BALCOM_SLFIHB, "
         g_str_Parame = g_str_Parame & "BALCOM_FECCRE, "
         g_str_Parame = g_str_Parame & "BALCOM_USUCRE) "

         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_PerMes & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_PerAno & "', "
         g_str_Parame = g_str_Parame & p_TipMon & ", "

         If Mid(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 3, 1) <> 0 Then
            If p_TipMon = 1 Then
               If Mid(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 3, 1) = 2 Then
                  If (Left(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 1) = 4 Or Left(Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux), 1) = 5) Then
                     g_str_Parame = g_str_Parame & 1 & ", "
                  Else
                     g_str_Parame = g_str_Parame & 3 & ", "
                  End If
               Else
                  g_str_Parame = g_str_Parame & 1 & ", "
               End If
            ElseIf p_TipMon = 2 Then
               g_str_Parame = g_str_Parame & 2 & ", "
            End If
         Else
            g_str_Parame = g_str_Parame & 0 & ", "
         End If

         g_str_Parame = g_str_Parame & "'" & Trim(r_arr_BalCom(r_int_NroIte).BalCom_CtaAux) & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_DesCue & "', "
         g_str_Parame = g_str_Parame & r_arr_BalCom(r_int_NroIte).BalCom_ImpDeb & ", "
         g_str_Parame = g_str_Parame & r_arr_BalCom(r_int_NroIte).BalCom_ImpHab & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & 0 & ", "
         g_str_Parame = g_str_Parame & "'" & l_str_FecImp & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "') "

         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
      End If
   Next r_int_NroIte

End Sub

Private Sub fs_SalFin(ByVal p_TipMon As Integer)

   Dim r_int_NroIte        As Integer

   Dim r_dbl_SalFin        As Double
   Dim r_dbl_SalIni        As Double
   Dim r_arr_BalCom()      As modtac_tpo_BalCom

   ReDim r_arr_BalCom(0)

   g_str_Parame = "SELECT * FROM CTB_BALCOM WHERE "
   g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_TIPMON = " & p_TipMon & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " "
   g_str_Parame = g_str_Parame & "ORDER BY BALCOM_TIPMON ASC, BALCOM_CUENTA ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then

      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF

         'r_dbl_SalIni = g_rst_Princi!BALCOM_SLINDB - g_rst_Princi!BALCOM_SLINHB
         'r_dbl_SalFin = r_dbl_SalIni + (g_rst_Princi!BalCom_ImpDeb - g_rst_Princi!BalCom_ImpHab)

         'If r_dbl_SalFin > 0 Then
         '   r_dbl_SalFin = r_dbl_SalFin
         'Else
         '   r_dbl_SalFin = r_dbl_SalFin * -1
         'End If

         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "UPDATE CTB_BALCOM SET "

         'If g_rst_Princi!BalCom_ImpDeb > g_rst_Princi!BalCom_ImpHab Then
         '   g_str_Parame = g_str_Parame & "BALCOM_SLFIDB = " & r_dbl_SalFin & ", "
         '   g_str_Parame = g_str_Parame & "BALCOM_SLFIHB = " & 0 & " "
         'ElseIf g_rst_Princi!BalCom_ImpDeb < g_rst_Princi!BalCom_ImpHab Then
         '   g_str_Parame = g_str_Parame & "BALCOM_SLFIDB = " & 0 & ", "
         '   g_str_Parame = g_str_Parame & "BALCOM_SLFIHB = " & r_dbl_SalFin & " "
         'Else
         '   If g_rst_Princi!BalCom_ImpDeb <> 0 And g_rst_Princi!BalCom_ImpHab <> 0 Then
         '      g_str_Parame = g_str_Parame & "BALCOM_SLFIDB = " & 0 & ", "
         '      g_str_Parame = g_str_Parame & "BALCOM_SLFIHB = " & 0 & " "
         '   ElseIf g_rst_Princi!BalCom_ImpDeb = 0 And g_rst_Princi!BalCom_ImpHab = 0 Then
         '      If g_rst_Princi!BALCOM_SLINDB > g_rst_Princi!BALCOM_SLINHB Then
         '         g_str_Parame = g_str_Parame & "BALCOM_SLFIDB = " & r_dbl_SalFin & ", "
         '         g_str_Parame = g_str_Parame & "BALCOM_SLFIHB = " & 0 & " "
         '     ElseIf g_rst_Princi!BALCOM_SLINDB < g_rst_Princi!BALCOM_SLINHB Then
         '         g_str_Parame = g_str_Parame & "BALCOM_SLFIDB = " & 0 & ", "
         '         g_str_Parame = g_str_Parame & "BALCOM_SLFIHB = " & r_dbl_SalFin & " "
         '      Else
         '         g_str_Parame = g_str_Parame & "BALCOM_SLFIDB = " & 0 & ", "
         '         g_str_Parame = g_str_Parame & "BALCOM_SLFIHB = " & 0 & " "
         '      End If
         '   End If
         'End If

         If Left((Trim(g_rst_Princi!BALCOM_CUENTA)), 1) = 1 Or Left((Trim(g_rst_Princi!BALCOM_CUENTA)), 1) = 4 Or Left((Trim(g_rst_Princi!BALCOM_CUENTA)), 1) = 6 Or Left((Trim(g_rst_Princi!BALCOM_CUENTA)), 1) = 7 Or Left((Trim(g_rst_Princi!BALCOM_CUENTA)), 1) = 8 Then

            If g_rst_Princi!BALCOM_SLINDB - g_rst_Princi!BALCOM_SLINHB + g_rst_Princi!BalCom_ImpDeb - g_rst_Princi!BalCom_ImpHab > 0 Then
               g_str_Parame = g_str_Parame & "BALCOM_SLFIDB = " & g_rst_Princi!BALCOM_SLINDB - g_rst_Princi!BALCOM_SLINHB + g_rst_Princi!BalCom_ImpDeb - g_rst_Princi!BalCom_ImpHab & " "
            Else
               g_str_Parame = g_str_Parame & "BALCOM_SLFIHB = " & (g_rst_Princi!BALCOM_SLINDB - g_rst_Princi!BALCOM_SLINHB + g_rst_Princi!BalCom_ImpDeb - g_rst_Princi!BalCom_ImpHab) * -1 & " "
            End If

         ElseIf Left((Trim(g_rst_Princi!BALCOM_CUENTA)), 1) = 2 Or Left((Trim(g_rst_Princi!BALCOM_CUENTA)), 1) = 3 Or Left((Trim(g_rst_Princi!BALCOM_CUENTA)), 1) = 5 Then

            If g_rst_Princi!BALCOM_SLINHB - g_rst_Princi!BALCOM_SLINDB + g_rst_Princi!BalCom_ImpHab - g_rst_Princi!BalCom_ImpDeb > 0 Then
               g_str_Parame = g_str_Parame & "BALCOM_SLFIHB = " & g_rst_Princi!BALCOM_SLINHB - g_rst_Princi!BALCOM_SLINDB + g_rst_Princi!BalCom_ImpHab - g_rst_Princi!BalCom_ImpDeb & " "
            Else
               g_str_Parame = g_str_Parame & "BALCOM_SLFIDB = " & (g_rst_Princi!BALCOM_SLINHB - g_rst_Princi!BALCOM_SLINDB + g_rst_Princi!BalCom_ImpHab - g_rst_Princi!BalCom_ImpDeb) * -1 & " "
            End If

         End If


         g_str_Parame = g_str_Parame & "WHERE "
         g_str_Parame = g_str_Parame & "BALCOM_CODEMP = '" & l_str_CodEmp & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_FECCRE = '" & l_str_FecImp & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_TIPMON = '" & p_TipMon & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_USUCRE = '" & modgen_g_str_CodUsu & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_PERMES = '" & l_str_PerMes & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_PERANO = '" & l_str_PerAno & "' AND "
         g_str_Parame = g_str_Parame & "BALCOM_CUENTA = '" & Trim(g_rst_Princi!BALCOM_CUENTA) & "' "

         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If

         g_rst_Princi.MoveNext

      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Sub

Private Sub fs_DepSal()
   
   g_str_Parame = "DELETE FROM CTB_BALCOM WHERE "
   g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "(BALCOM_SLINHB = 0 AND "
   g_str_Parame = g_str_Parame & "BALCOM_SLINDB = 0 AND "
   g_str_Parame = g_str_Parame & "BALCOM_IMPDEB = 0 AND "
   g_str_Parame = g_str_Parame & "BALCOM_IMPHAB = 0 AND "
   g_str_Parame = g_str_Parame & "BALCOM_SLFIDB = 0 AND "
   g_str_Parame = g_str_Parame & "BALCOM_SLFIHB = 0 )AND "
   g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
End Sub

Public Function ff_TotMon_Sol1() As Integer

   ff_TotMon_Sol1 = 0

   g_str_Parame = "SELECT CNTA_CTBL, FLAG_DEBHAB, SUM(IMP_MOVSOL) AS IMPSOL, SUM(IMP_MOVDOL) AS IMPDOL FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & IIf(l_str_PerMes = 13, l_str_PerMes - 1, l_str_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO <> 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL, FLAG_DEBHAB "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL, FLAG_DEBHAB ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF

         ff_TotMon_Sol1 = ff_TotMon_Sol1 + 1

         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Function

Public Function ff_TotMon_Sol2() As Integer

   ff_TotMon_Sol2 = 0

   g_str_Parame = "SELECT DISTINCT(CNTA_CTBL) FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & IIf(l_str_PerMes = 13, l_str_PerMes - 1, l_str_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO <> 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF

         ff_TotMon_Sol2 = ff_TotMon_Sol2 + 1

         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Function

Public Function ff_TotMon_Dol1() As Integer

   ff_TotMon_Dol1 = 0

   g_str_Parame = "SELECT CNTA_CTBL, FLAG_DEBHAB, SUM(IMP_MOVSOL) AS IMPSOL, SUM(IMP_MOVDOL) AS IMPDOL FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & IIf(l_str_PerMes = 13, l_str_PerMes - 1, l_str_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,3,1)  = 2 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 4 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 5 AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO <> 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL, FLAG_DEBHAB "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL, FLAG_DEBHAB ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF

         ff_TotMon_Dol1 = ff_TotMon_Dol1 + 1

         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Function

Public Function ff_TotMon_Dol2() As Integer

   ff_TotMon_Dol2 = 0

   g_str_Parame = "SELECT DISTINCT(CNTA_CTBL) FROM CNTBL_ASIENTO_DET WHERE "
   g_str_Parame = g_str_Parame & "ANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "MES = " & IIf(l_str_PerMes = 13, l_str_PerMes - 1, l_str_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,3,1)  = 2 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 4 AND "
   g_str_Parame = g_str_Parame & "SUBSTR(CNTA_CTBL,1,1) <> 5 AND "
   g_str_Parame = g_str_Parame & "NRO_LIBRO <> 5 "
   g_str_Parame = g_str_Parame & "GROUP BY CNTA_CTBL "
   g_str_Parame = g_str_Parame & "ORDER BY CNTA_CTBL ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF

         ff_TotMon_Dol2 = ff_TotMon_Dol2 + 1

         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Function

Public Function ff_Buscar() As Integer
   ff_Buscar = 0

   g_str_Parame = "SELECT COUNT(*) AS TOTREG FROM CTB_BALCOM WHERE "
   g_str_Parame = g_str_Parame & "BALCOM_PERANO = " & l_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_PERMES = " & l_str_PerMes & " AND "
   g_str_Parame = g_str_Parame & "BALCOM_CODEMP = " & l_str_CodEmp & " "

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

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub
