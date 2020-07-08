VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptSun_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   10365
   ClientTop       =   2835
   ClientWidth     =   5940
   Icon            =   "GesCtb_frm_856.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2475
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5955
      _Version        =   65536
      _ExtentX        =   10504
      _ExtentY        =   4366
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
            TabIndex        =   7
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
            TabIndex        =   8
            Top             =   315
            Width           =   4305
            _Version        =   65536
            _ExtentX        =   7594
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "F08.1. - Registro de Compras"
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
            Picture         =   "GesCtb_frm_856.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   915
         Left            =   30
         TabIndex        =   10
         Top             =   1470
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
         _ExtentY        =   1614
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
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   2500
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1080
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
            Caption         =   "Año :"
            Height          =   285
            Left            =   210
            TabIndex        =   12
            Top             =   540
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo :"
            Height          =   315
            Left            =   210
            TabIndex        =   11
            Top             =   180
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   9
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
         Begin VB.CommandButton cmd_Archivo 
            Height          =   585
            Left            =   660
            Picture         =   "GesCtb_frm_856.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Generar archivo texto"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5190
            Picture         =   "GesCtb_frm_856.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   60
            Picture         =   "GesCtb_frm_856.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_RptSun_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_bol_FlgReg     As Boolean
Dim r_int_PerMes     As Integer
Dim r_int_PerAno     As Integer

Private Sub cmd_ExpExc_Click()
   If cmb_CodMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodMes)
      Exit Sub
   End If
   
   If ipp_PerAno.Text = "" Then
      MsgBox "Debe seleccionar el Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   r_int_PerMes = CInt(cmb_CodMes.ItemData(cmb_CodMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)
   'Call fs_GenExcPLE
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Archivo_Click()
Dim r_int_NroCor     As Integer
Dim r_str_NomRes     As String
Dim r_str_NumRuc     As String
Dim r_str_DetGlo     As String
Dim r_dbl_TipCam     As Double
Dim r_int_NumRes     As Integer
Dim r_rst_Total      As ADODB.Recordset
Dim r_str_Nombre     As String

   If cmb_CodMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = "" Then
      MsgBox "Debe seleccionar el Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If

   'Verifica que exista ruta
   If Dir$(moddat_g_str_RutLoc, vbDirectory) = "" Then
      MsgBox "Debe crear el siguente directorio " & moddat_g_str_RutLoc, vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   Screen.MousePointer = 11
   r_int_PerMes = CInt(cmb_CodMes.ItemData(cmb_CodMes.ListIndex))
   r_int_PerAno = CInt(ipp_PerAno.Text)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUBSTR(TO_CHAR(FECHA_CNTBL,'DDMMYYYY'),5,4) || SUBSTR(TO_CHAR(FECHA_CNTBL,'DDMMYYYY'),3,2) || '00' AS CAMPO_01,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_02,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_03,"
   g_str_Parame = g_str_Parame & "  FECHA_CNTBL                            AS CAMPO_04, "
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_05,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,13,2)                 AS CAMPO_06,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,16,4)                 AS CAMPO_07,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_08,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,21,7)                 AS CAMPO_09,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_10,"
   g_str_Parame = g_str_Parame & "  '6'                                    AS CAMPO_11,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,1,11)                 AS CAMPO_12,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,29,31)                AS CAMPO_13,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_14,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_15,"
   g_str_Parame = g_str_Parame & "  ROUND((IMP_MOVSOL/1.18),2)             AS CAMPO_16,"
   g_str_Parame = g_str_Parame & "  IMP_MOVSOL-ROUND((IMP_MOVSOL/1.18),2)  AS CAMPO_17,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_18,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_19,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_20,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_21,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_22,"
   g_str_Parame = g_str_Parame & "  ROUND(IMP_MOVSOL,2)                    AS CAMPO_23,"
   g_str_Parame = g_str_Parame & "  0.000                                  AS CAMPO_24,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_25,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_26,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_27,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_28,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_29,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_30,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_31,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_32,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_33,"
   g_str_Parame = g_str_Parame & "  1                                      AS CAMPO_34"
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO_DET"
   g_str_Parame = g_str_Parame & " WHERE ANO = '" & ipp_PerAno.Text & "'"
   g_str_Parame = g_str_Parame & "   AND MES = '" & cmb_CodMes.ListIndex + 1 & "'"
   g_str_Parame = g_str_Parame & "   AND NRO_LIBRO = 15"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      MsgBox "No hay datos.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Creando Archivo
   r_str_NomRes = moddat_g_str_RutLoc & "\LE20511904162" & r_int_PerAno & Format(r_int_PerMes, "00") & "00" & "08" & "0100001111" & ".txt"
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes

   r_int_NroCor = 1

   Do While Not g_rst_Princi.EOF
      r_str_NumRuc = Val(Mid(Trim(g_rst_Princi!CAMPO_12), 1, 11))
      
      If InStr(g_rst_Princi!CAMPO_13, "/") - 1 > 0 Then
         r_str_Nombre = Mid(Trim(g_rst_Princi!CAMPO_13), 1, InStr(g_rst_Princi!CAMPO_13, "/") - 1)
      Else
         r_str_Nombre = Mid(IIf(IsNull(g_rst_Princi!CAMPO_13), "", Trim(g_rst_Princi!CAMPO_13)), 1)
      End If
      
      If gf_Valida_RUC(r_str_NumRuc, Mid(r_str_NumRuc, 11, 1)) Then
         Print #1, IIf(IsNull(g_rst_Princi!CAMPO_01), "", Trim(g_rst_Princi!CAMPO_01)); "|"; Format(r_int_NroCor, "000"); "|"; "M" & Format(r_int_NroCor, "000"); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_04), "", Trim(g_rst_Princi!CAMPO_04)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_04), "", IIf(g_rst_Princi!CAMPO_06 = "01", "", Trim(g_rst_Princi!CAMPO_04))); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_06), "", Trim(g_rst_Princi!CAMPO_06)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_07), "", Trim(g_rst_Princi!CAMPO_07)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_08), "", Trim(g_rst_Princi!CAMPO_08)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_09), "", Trim(g_rst_Princi!CAMPO_09)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_10), "", Trim(g_rst_Princi!CAMPO_10)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_11), "", Trim(g_rst_Princi!CAMPO_11)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_12), "", Trim(g_rst_Princi!CAMPO_12)); "|"; IIf(IsNull(r_str_Nombre), "", Trim(r_str_Nombre)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_14), "", Trim(g_rst_Princi!CAMPO_14)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_15), "", Trim(g_rst_Princi!CAMPO_15)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_16), "", Trim(g_rst_Princi!CAMPO_16)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_17), "", Trim(g_rst_Princi!CAMPO_17)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_18), "", Trim(g_rst_Princi!CAMPO_18)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_19), "", Trim(g_rst_Princi!CAMPO_19)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_20), "", Trim(g_rst_Princi!CAMPO_20)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_21), "", Trim(g_rst_Princi!CAMPO_21)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_22), "", Trim(g_rst_Princi!CAMPO_22)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_23), "", Trim(g_rst_Princi!CAMPO_23)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_24), "", Format(g_rst_Princi!CAMPO_24, "0.000")); "|"; IIf(IsNull(g_rst_Princi!CAMPO_25), "", Trim(g_rst_Princi!CAMPO_25)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_26), "", Trim(g_rst_Princi!CAMPO_26)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_27), "", Trim(g_rst_Princi!CAMPO_27)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_28), "", Trim(g_rst_Princi!CAMPO_28)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_29), "", Trim(g_rst_Princi!CAMPO_29)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_30), "", Trim(g_rst_Princi!CAMPO_30)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_31), "", Trim(g_rst_Princi!CAMPO_31)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_32), "", Trim(g_rst_Princi!CAMPO_32)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_33), "", Trim(g_rst_Princi!CAMPO_33)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_34), "", Trim(g_rst_Princi!CAMPO_34)); "|"
                  
         
         r_int_NroCor = r_int_NroCor + 1
      End If
      g_rst_Princi.MoveNext
      DoEvents
   Loop
            
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Close #1
   
   Screen.MousePointer = 0
   MsgBox "El archivo ha sido creado: " & Trim(r_str_NomRes), vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_CodMes, 1, "033")
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_CodMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
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

Private Sub fs_GenExc()
Dim r_rst_Total      As ADODB.Recordset
Dim r_obj_Excel      As Excel.Application
Dim r_int_CntFil     As Integer
Dim r_int_NroCor     As Integer
Dim r_str_NumRuc     As String

   r_int_NroCor = 1
   r_int_CntFil = 13
   
   '****************
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUBSTR(TO_CHAR(FECHA_CNTBL,'DDMMYYYY'),5,4) || SUBSTR(TO_CHAR(FECHA_CNTBL,'DDMMYYYY'),3,2) || '00' AS CAMPO_01,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_02,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_03,"
   g_str_Parame = g_str_Parame & "  FECHA_CNTBL                            AS CAMPO_04, "
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_05,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,13,2)                 AS CAMPO_06,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,16,4)                 AS CAMPO_07,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,21,7)                 AS CAMPO_08,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_09,"
   g_str_Parame = g_str_Parame & "  '06'                                   AS CAMPO_10,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,1,11)                 AS CAMPO_11,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,29,31)                AS CAMPO_12,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_13,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_14,"
   g_str_Parame = g_str_Parame & "  ROUND((IMP_MOVSOL/1.18),2)             AS CAMPO_15,"
   g_str_Parame = g_str_Parame & "  IMP_MOVSOL-ROUND((IMP_MOVSOL/1.18),2)  AS CAMPO_16,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_17,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_18,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_19,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_20,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_21,"
   g_str_Parame = g_str_Parame & "  ROUND(IMP_MOVSOL,2)                    AS CAMPO_22,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_23,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_24,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_25,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_26,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_27,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_28,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_29,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_30,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_31,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_32,"
   g_str_Parame = g_str_Parame & "  1                                      AS CAMPO_33"
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO_DET"
   g_str_Parame = g_str_Parame & " WHERE ANO = '" & ipp_PerAno.Text & "'"
   g_str_Parame = g_str_Parame & "   AND MES = '" & cmb_CodMes.ListIndex + 1 & "'"
   g_str_Parame = g_str_Parame & "   AND NRO_LIBRO = 15"
   g_str_Parame = g_str_Parame & "   AND FLAG_DEBHAB = 'H'"
   
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
   r_obj_Excel.Sheets(1).Name = "REGISTRO DE COMPRAS"
   
   With r_obj_Excel.Sheets(1)
      .Cells(2, 1) = "FORMATO 08.1: REGISTRO DE COMPRAS"
      .Cells(4, 1) = "PERIODO:  " & Trim(cmb_CodMes.Text) & "  " & Trim(ipp_PerAno.Text)
      .Cells(5, 1) = "RUC    :  20511904162"
      .Cells(6, 1) = "RAZON SOCIAL:  EDPYME MICASITA S.A."
      .Cells(7, 1) = "OFICINA:  PRINCIPAL"
      .Rows(12).RowHeight = 30
      
      .Range(.Cells(10, 1), .Cells(12, 1)).Merge
      .Range(.Cells(10, 1), .Cells(12, 1)).WrapText = True
      .Range(.Cells(10, 1), .Cells(12, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 1), .Cells(12, 1)).ColumnWidth = 10
      .Range(.Cells(10, 2), .Cells(12, 2)).Merge
      .Range(.Cells(10, 2), .Cells(12, 2)).WrapText = True
      .Range(.Cells(10, 2), .Cells(12, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 2), .Cells(12, 2)).ColumnWidth = 13.71
      .Range(.Cells(10, 3), .Cells(12, 3)).Merge
      .Range(.Cells(10, 3), .Cells(12, 3)).WrapText = True
      .Range(.Cells(10, 3), .Cells(12, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 3), .Cells(12, 3)).ColumnWidth = 10
      .Range(.Cells(10, 4), .Cells(12, 4)).Merge
      .Range(.Cells(10, 4), .Cells(12, 4)).WrapText = True
      .Range(.Cells(10, 4), .Cells(12, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 4), .Cells(12, 4)).ColumnWidth = 11
      .Range(.Cells(10, 5), .Cells(12, 5)).Merge
      .Range(.Cells(10, 5), .Cells(12, 5)).WrapText = True
      .Range(.Cells(10, 5), .Cells(12, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 5), .Cells(12, 5)).ColumnWidth = 13
      
      'concatenar
      .Range(.Cells(10, 6), .Cells(10, 9)).Merge
      .Range(.Cells(10, 6), .Cells(10, 9)).WrapText = True
      .Range(.Cells(10, 6), .Cells(10, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 6), .Cells(12, 6)).Merge
      .Range(.Cells(11, 6), .Cells(12, 6)).WrapText = True
      .Range(.Cells(11, 6), .Cells(12, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 6), .Cells(12, 6)).ColumnWidth = 10
      .Range(.Cells(11, 7), .Cells(12, 7)).Merge
      .Range(.Cells(11, 7), .Cells(12, 7)).WrapText = True
      .Range(.Cells(11, 7), .Cells(12, 7)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 7), .Cells(12, 7)).ColumnWidth = 11
      .Range(.Cells(11, 8), .Cells(12, 8)).Merge
      .Range(.Cells(11, 8), .Cells(12, 8)).WrapText = True
      .Range(.Cells(11, 8), .Cells(12, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 8), .Cells(12, 8)).ColumnWidth = 12
      .Range(.Cells(11, 9), .Cells(12, 9)).Merge
      .Range(.Cells(11, 9), .Cells(12, 9)).WrapText = True
      .Range(.Cells(11, 9), .Cells(12, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 9), .Cells(12, 9)).ColumnWidth = 12
      
      'concatenar
      .Range(.Cells(10, 10), .Cells(10, 12)).Merge
      .Range(.Cells(10, 10), .Cells(10, 12)).WrapText = True
      .Range(.Cells(10, 10), .Cells(10, 12)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("J").ColumnWidth = 10
      .Cells(12, 10).WrapText = True
      .Cells(12, 10).HorizontalAlignment = xlHAlignCenter
      
      .Columns("K").ColumnWidth = 12
      .Columns("L").ColumnWidth = 31
      .Cells(12, 11).WrapText = True
      .Cells(12, 11).HorizontalAlignment = xlHAlignCenter
      
      'concatenar
      .Range(.Cells(11, 10), .Cells(11, 11)).Merge
      .Range(.Cells(11, 10), .Cells(11, 11)).WrapText = True
      .Range(.Cells(11, 10), .Cells(11, 11)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 12), .Cells(12, 12)).Merge
      .Range(.Cells(11, 12), .Cells(12, 12)).WrapText = True
      .Range(.Cells(11, 12), .Cells(12, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 12), .Cells(12, 12)).ColumnWidth = 41

      .Range(.Cells(10, 13), .Cells(10, 14)).Merge
      .Range(.Cells(10, 13), .Cells(10, 14)).WrapText = True
      .Range(.Cells(10, 13), .Cells(10, 14)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(11, 13), .Cells(12, 13)).Merge
      .Range(.Cells(11, 13), .Cells(12, 13)).WrapText = True
      .Range(.Cells(11, 13), .Cells(12, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 13), .Cells(12, 13)).ColumnWidth = 14
      .Range(.Cells(11, 14), .Cells(12, 14)).Merge
      .Range(.Cells(11, 14), .Cells(12, 14)).WrapText = True
      .Range(.Cells(11, 14), .Cells(12, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 14), .Cells(12, 14)).ColumnWidth = 16
      
      .Range(.Cells(10, 15), .Cells(12, 15)).Merge
      .Range(.Cells(10, 15), .Cells(12, 15)).WrapText = True
      .Range(.Cells(10, 15), .Cells(12, 15)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 15), .Cells(12, 15)).ColumnWidth = 12
      
      .Range(.Cells(10, 16), .Cells(12, 16)).Merge
      .Range(.Cells(10, 16), .Cells(12, 16)).WrapText = True
      .Range(.Cells(10, 16), .Cells(12, 16)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 16), .Cells(12, 16)).ColumnWidth = 12

      .Range(.Cells(10, 17), .Cells(12, 17)).Merge
      .Range(.Cells(10, 17), .Cells(12, 17)).WrapText = True
      .Range(.Cells(10, 17), .Cells(12, 17)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 17), .Cells(12, 17)).ColumnWidth = 12
      
      .Range(.Cells(10, 18), .Cells(12, 18)).Merge
      .Range(.Cells(10, 18), .Cells(12, 18)).WrapText = True
      .Range(.Cells(10, 18), .Cells(12, 18)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 18), .Cells(12, 18)).ColumnWidth = 12
      
      .Range(.Cells(10, 19), .Cells(10, 20)).Merge
      .Range(.Cells(10, 19), .Cells(10, 20)).WrapText = True
      .Range(.Cells(10, 19), .Cells(10, 20)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 19), .Cells(12, 19)).Merge
      .Range(.Cells(11, 19), .Cells(12, 19)).WrapText = True
      .Range(.Cells(11, 19), .Cells(12, 19)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 19), .Cells(12, 19)).ColumnWidth = 13
      .Range(.Cells(11, 20), .Cells(12, 20)).Merge
      .Range(.Cells(11, 20), .Cells(12, 20)).WrapText = True
      .Range(.Cells(11, 20), .Cells(12, 20)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(11, 20), .Cells(12, 20)).ColumnWidth = 13
      
      .Range(.Cells(10, 21), .Cells(12, 21)).Merge
      .Range(.Cells(10, 21), .Cells(12, 21)).WrapText = True
      .Range(.Cells(10, 21), .Cells(12, 21)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 21), .Cells(12, 21)).ColumnWidth = 8
      .Range(.Cells(10, 22), .Cells(12, 22)).Merge
      .Range(.Cells(10, 22), .Cells(12, 22)).WrapText = True
      .Range(.Cells(10, 22), .Cells(12, 22)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 22), .Cells(12, 22)).ColumnWidth = 8
      
      .Cells(10, 1) = "PERIODO"
      .Cells(10, 2) = "NUMERO CORRELATIVO"
      .Cells(10, 3) = "NUMERO CUO"
      .Cells(10, 4) = "FECHA DE EMISION"
      .Cells(10, 5) = "FECHA DE VENCIMIENTO Y/O PAGO"
      .Cells(10, 6) = "COMPROBANTE DE PAGO O DOCUMENTO"
      .Cells(11, 6) = "TIPO (TABLA 10)"
      .Cells(11, 7) = "NRO. SERIE"
      .Cells(11, 8) = "NUMERO INICIAL"
      .Cells(11, 9) = "NUMERO FINAL"
      .Cells(10, 10) = "INFORMACION DEL PROVEEDOR"
      .Cells(11, 10) = "DOCUMENTO IDENTIDAD"
      .Cells(12, 10) = "TIPO (TABLA 2)"
      .Cells(12, 11) = "NUMERO"
      .Cells(11, 12) = "APELLIDOS Y NOMBRES DENOMINACION O RAZON SOCIAL"
      .Cells(10, 13) = "GRAVADAS DESTINADAS A VENTAS NO GRAVADAS"
      .Cells(11, 13) = "BASE IMPONIBLE"
      .Cells(11, 14) = "IGV"
      .Cells(10, 15) = "NO GRAVADAS"
      .Cells(10, 16) = "ISC"
      .Cells(10, 17) = "OTROS TRIBUTOS Y CARGOS"
      .Cells(10, 18) = "IMPORTE TOTAL"
      .Cells(10, 19) = "CONSTANCIA DE DEPOSITO DE DETRACCION"
      .Cells(11, 19) = "NUMERO"
      .Cells(11, 20) = "FECHA"
      .Cells(10, 21) = "TIPO DE CAMBIO"
      .Cells(10, 22) = "ESTADO"
      
      .Range(.Cells(1, 1), .Cells(12, 22)).Font.Bold = True
      .Range(.Cells(10, 1), .Cells(12, 22)).Interior.Color = RGB(146, 208, 80)
      
      .Columns("M").NumberFormat = "#,###,##0.00"
      .Columns("N").NumberFormat = "#,###,##0.00"
      .Columns("O").NumberFormat = "#,###,##0.00"
      .Columns("P").NumberFormat = "#,###,##0.00"
      .Columns("Q").NumberFormat = "#,###,##0.00"
      .Columns("R").NumberFormat = "#,###,##0.00"
      .Columns("U").NumberFormat = "#,###,##0.000"
      
      'r_obj_Excel.Visible = True
      .Range(.Cells(10, 1), .Cells(12, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      Do Until r_int_NroCor > 22
         .Range(.Cells(12, r_int_NroCor), .Cells(12, r_int_NroCor)).VerticalAlignment = xlCenter
         .Range(.Cells(10, r_int_NroCor), .Cells(10, r_int_NroCor)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(12, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(10, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(10, r_int_NroCor), .Cells(10, r_int_NroCor)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(12, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(10, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeRight).LineStyle = xlContinuous
         r_int_NroCor = r_int_NroCor + 1
      Loop
      
      .Range(.Cells(10, 6), .Cells(10, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(11, 10), .Cells(11, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(10, 10), .Cells(10, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(10, 13), .Cells(10, 22)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroCor = 1
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_str_NumRuc = Val(Mid(Trim(g_rst_Princi!CAMPO_11), 1, 11))
         
         If gf_Valida_RUC(r_str_NumRuc, Mid(r_str_NumRuc, 11, 1)) Then
            .Cells(r_int_CntFil, 1) = g_rst_Princi!CAMPO_01
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 1)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 2) = "'" & Format(r_int_NroCor, "000")
            .Range(.Cells(r_int_CntFil, 2), .Cells(r_int_CntFil, 2)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 3) = "M" & Format(r_int_NroCor, "000")
            .Range(.Cells(r_int_CntFil, 3), .Cells(r_int_CntFil, 3)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 4) = "'" & Trim(g_rst_Princi!CAMPO_04)
            .Range(.Cells(r_int_CntFil, 4), .Cells(r_int_CntFil, 4)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 5) = IIf(IsNull(g_rst_Princi!CAMPO_05), "", g_rst_Princi!CAMPO_05)
            
            .Cells(r_int_CntFil, 6) = IIf(IsNull(g_rst_Princi!CAMPO_06), "", g_rst_Princi!CAMPO_06)
            .Range(.Cells(r_int_CntFil, 6), .Cells(r_int_CntFil, 6)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 7) = IIf(IsNull(g_rst_Princi!CAMPO_07), "", g_rst_Princi!CAMPO_07)
            .Range(.Cells(r_int_CntFil, 7), .Cells(r_int_CntFil, 7)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 8) = IIf(IsNull(g_rst_Princi!CAMPO_08), "", g_rst_Princi!CAMPO_08)
            .Range(.Cells(r_int_CntFil, 8), .Cells(r_int_CntFil, 8)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 9) = IIf(IsNull(g_rst_Princi!CAMPO_09), "", g_rst_Princi!CAMPO_09)
            
            .Cells(r_int_CntFil, 10) = IIf(IsNull(g_rst_Princi!CAMPO_10), "", g_rst_Princi!CAMPO_10)
            .Range(.Cells(r_int_CntFil, 10), .Cells(r_int_CntFil, 10)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 11) = IIf(IsNull(g_rst_Princi!CAMPO_11), "", g_rst_Princi!CAMPO_11)
            .Range(.Cells(r_int_CntFil, 11), .Cells(r_int_CntFil, 11)).HorizontalAlignment = xlHAlignCenter
            
            
            If InStr(g_rst_Princi!CAMPO_12, "/") - 1 > 0 Then
               .Cells(r_int_CntFil, 12) = Mid(Trim(g_rst_Princi!CAMPO_12), 1, InStr(g_rst_Princi!CAMPO_12, "/") - 1)
            Else
               .Cells(r_int_CntFil, 12) = Mid(Trim(g_rst_Princi!CAMPO_12), 1)
            End If

            .Cells(r_int_CntFil, 13) = IIf(IsNull(g_rst_Princi!CAMPO_15), "", g_rst_Princi!CAMPO_15)
            .Cells(r_int_CntFil, 14) = IIf(IsNull(g_rst_Princi!CAMPO_16), "", g_rst_Princi!CAMPO_16)
            .Cells(r_int_CntFil, 15) = IIf(IsNull(g_rst_Princi!CAMPO_17), "0", g_rst_Princi!CAMPO_17)
            .Cells(r_int_CntFil, 16) = IIf(IsNull(g_rst_Princi!CAMPO_18), "", g_rst_Princi!CAMPO_18)
            .Cells(r_int_CntFil, 17) = IIf(IsNull(g_rst_Princi!CAMPO_19), "", g_rst_Princi!CAMPO_19)
            .Cells(r_int_CntFil, 18) = IIf(IsNull(g_rst_Princi!CAMPO_22), "", g_rst_Princi!CAMPO_22)
            .Cells(r_int_CntFil, 19) = IIf(IsNull(g_rst_Princi!CAMPO_24), "", g_rst_Princi!CAMPO_24)
            .Cells(r_int_CntFil, 20) = IIf(IsNull(g_rst_Princi!CAMPO_25), "", g_rst_Princi!CAMPO_25)
            .Cells(r_int_CntFil, 21) = IIf(IsNull(g_rst_Princi!CAMPO_26), "0", g_rst_Princi!CAMPO_26)
            .Cells(r_int_CntFil, 22) = "1"
            
            .Range(.Cells(r_int_CntFil, 22), .Cells(r_int_CntFil, 22)).HorizontalAlignment = xlHAlignCenter
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 22)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 22)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 22)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 22)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 22)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
            r_int_NroCor = r_int_NroCor + 1
            r_int_CntFil = r_int_CntFil + 1
         End If
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExcPLE()
Dim r_rst_Total      As ADODB.Recordset
Dim r_obj_Excel      As Excel.Application
Dim r_int_CntFil     As Integer
Dim r_int_NroCor     As Integer
Dim r_str_NumRuc     As String

   r_int_NroCor = 1
   r_int_CntFil = 13
   
   '****************
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT '20' || SUBSTR(FECHA_CNTBL,7,4) || SUBSTR(FECHA_CNTBL,4,2) || '00' AS CAMPO_01,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_02,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_03,"
   g_str_Parame = g_str_Parame & "  FECHA_CNTBL                            AS CAMPO_04, "
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_05,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,13,2)                 AS CAMPO_06,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,16,4)                 AS CAMPO_07,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,21,7)                 AS CAMPO_08,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_09,"
   g_str_Parame = g_str_Parame & "  '06'                                   AS CAMPO_10,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,1,11)                 AS CAMPO_11,"
   g_str_Parame = g_str_Parame & "  SUBSTR(DET_GLOSA,29,31)                AS CAMPO_12,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_13,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_14,"
   g_str_Parame = g_str_Parame & "  ROUND((IMP_MOVSOL/1.18),2)             AS CAMPO_15,"
   g_str_Parame = g_str_Parame & "  IMP_MOVSOL-ROUND((IMP_MOVSOL/1.18),2)  AS CAMPO_16,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_17,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_18,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_19,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_20,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_21,"
   g_str_Parame = g_str_Parame & "  ROUND(IMP_MOVSOL,2)                    AS CAMPO_22,"
   g_str_Parame = g_str_Parame & "  0                                      AS CAMPO_23,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_24,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_25,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_26,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_27,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_28,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_29,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_30,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_31,"
   g_str_Parame = g_str_Parame & "  ''                                     AS CAMPO_32,"
   g_str_Parame = g_str_Parame & "  1                                      AS CAMPO_33"
   g_str_Parame = g_str_Parame & "  FROM CNTBL_ASIENTO_DET"
   g_str_Parame = g_str_Parame & " WHERE ANO = '" & ipp_PerAno.Text & "'"
   g_str_Parame = g_str_Parame & "   AND MES = '" & cmb_CodMes.ListIndex + 1 & "'"
   g_str_Parame = g_str_Parame & "   AND NRO_LIBRO = 15"
 
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
   r_obj_Excel.Sheets(1).Name = "REGISTRO DE COMPRAS"
   
   With r_obj_Excel.Sheets(1)
      .Cells(2, 1) = "FORMATO 08.1: REGISTRO DE COMPRAS"
      .Cells(4, 1) = "PERIODO:  " & Trim(cmb_CodMes.Text) & "  " & Trim(ipp_PerAno.Text)
      .Cells(5, 1) = "RUC    :  20511904162"
      .Cells(6, 1) = "RAZON SOCIAL:  EDPYME MICASITA S.A."
      .Cells(7, 1) = "OFICINA:  PRINCIPAL"
      .Rows(12).RowHeight = 30
      
      .Range(.Cells(10, 1), .Cells(12, 1)).Merge
      .Range(.Cells(10, 1), .Cells(12, 1)).WrapText = True
      .Range(.Cells(10, 1), .Cells(12, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 1), .Cells(12, 1)).ColumnWidth = 10
      
      .Range(.Cells(10, 2), .Cells(12, 2)).Merge
      .Range(.Cells(10, 2), .Cells(12, 2)).WrapText = True
      .Range(.Cells(10, 2), .Cells(12, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 2), .Cells(12, 2)).ColumnWidth = 6
      
      .Range(.Cells(10, 3), .Cells(12, 3)).Merge
      .Range(.Cells(10, 3), .Cells(12, 3)).WrapText = True
      .Range(.Cells(10, 3), .Cells(12, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 3), .Cells(12, 3)).ColumnWidth = 9
      
      .Range(.Cells(10, 4), .Cells(12, 4)).Merge
      .Range(.Cells(10, 4), .Cells(12, 4)).WrapText = True
      .Range(.Cells(10, 4), .Cells(12, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 4), .Cells(12, 4)).ColumnWidth = 11
      
      .Range(.Cells(10, 5), .Cells(12, 5)).Merge
      .Range(.Cells(10, 5), .Cells(12, 5)).WrapText = True
      .Range(.Cells(10, 5), .Cells(12, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 5), .Cells(12, 5)).ColumnWidth = 11
      
      .Range(.Cells(10, 6), .Cells(12, 6)).Merge
      .Range(.Cells(10, 6), .Cells(12, 6)).WrapText = True
      .Range(.Cells(10, 6), .Cells(12, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 6), .Cells(12, 6)).ColumnWidth = 10
      
      .Range(.Cells(10, 7), .Cells(12, 7)).Merge
      .Range(.Cells(10, 7), .Cells(12, 7)).WrapText = True
      .Range(.Cells(10, 7), .Cells(12, 7)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 7), .Cells(12, 7)).ColumnWidth = 7
      
      .Range(.Cells(10, 8), .Cells(12, 8)).Merge
      .Range(.Cells(10, 8), .Cells(12, 8)).WrapText = True
      .Range(.Cells(10, 8), .Cells(12, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 8), .Cells(12, 8)).ColumnWidth = 9
      
      .Range(.Cells(10, 9), .Cells(12, 9)).Merge
      .Range(.Cells(10, 9), .Cells(12, 9)).WrapText = True
      .Range(.Cells(10, 9), .Cells(12, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 9), .Cells(12, 9)).ColumnWidth = 9
      
      .Range(.Cells(10, 10), .Cells(12, 10)).Merge
      .Range(.Cells(10, 10), .Cells(12, 10)).WrapText = True
      .Range(.Cells(10, 10), .Cells(12, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 10), .Cells(12, 10)).ColumnWidth = 9
            
      .Range(.Cells(10, 11), .Cells(12, 11)).Merge
      .Range(.Cells(10, 11), .Cells(12, 11)).WrapText = True
      .Range(.Cells(10, 11), .Cells(12, 11)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 11), .Cells(12, 11)).ColumnWidth = 14

      .Range(.Cells(10, 12), .Cells(12, 12)).Merge
      .Range(.Cells(10, 12), .Cells(12, 12)).WrapText = True
      .Range(.Cells(10, 12), .Cells(12, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 12), .Cells(12, 12)).ColumnWidth = 30

      .Range(.Cells(10, 12), .Cells(12, 12)).Merge
      .Range(.Cells(10, 12), .Cells(12, 12)).WrapText = True
      .Range(.Cells(10, 12), .Cells(12, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 12), .Cells(12, 12)).ColumnWidth = 30
      
      .Range(.Cells(10, 13), .Cells(12, 13)).Merge
      .Range(.Cells(10, 13), .Cells(12, 13)).WrapText = True
      .Range(.Cells(10, 13), .Cells(12, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 13), .Cells(12, 13)).ColumnWidth = 12
      
      .Range(.Cells(10, 14), .Cells(12, 14)).Merge
      .Range(.Cells(10, 14), .Cells(12, 14)).WrapText = True
      .Range(.Cells(10, 14), .Cells(12, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 14), .Cells(12, 14)).ColumnWidth = 12
      
      .Range(.Cells(10, 15), .Cells(12, 15)).Merge
      .Range(.Cells(10, 15), .Cells(12, 15)).WrapText = True
      .Range(.Cells(10, 15), .Cells(12, 15)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 15), .Cells(12, 15)).ColumnWidth = 12
      
      .Range(.Cells(10, 16), .Cells(12, 16)).Merge
      .Range(.Cells(10, 16), .Cells(12, 16)).WrapText = True
      .Range(.Cells(10, 16), .Cells(12, 16)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 16), .Cells(12, 16)).ColumnWidth = 12
      
      .Range(.Cells(10, 17), .Cells(12, 17)).Merge
      .Range(.Cells(10, 17), .Cells(12, 17)).WrapText = True
      .Range(.Cells(10, 17), .Cells(12, 17)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 17), .Cells(12, 17)).ColumnWidth = 12
      
      .Range(.Cells(10, 18), .Cells(12, 18)).Merge
      .Range(.Cells(10, 18), .Cells(12, 18)).WrapText = True
      .Range(.Cells(10, 18), .Cells(12, 18)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 18), .Cells(12, 18)).ColumnWidth = 12
      
      .Range(.Cells(10, 19), .Cells(12, 19)).Merge
      .Range(.Cells(10, 19), .Cells(12, 19)).WrapText = True
      .Range(.Cells(10, 19), .Cells(12, 19)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 19), .Cells(12, 19)).ColumnWidth = 12
      
      .Range(.Cells(10, 20), .Cells(12, 20)).Merge
      .Range(.Cells(10, 20), .Cells(12, 20)).WrapText = True
      .Range(.Cells(10, 20), .Cells(12, 20)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 20), .Cells(12, 20)).ColumnWidth = 12
      
      .Range(.Cells(10, 21), .Cells(12, 21)).Merge
      .Range(.Cells(10, 21), .Cells(12, 21)).WrapText = True
      .Range(.Cells(10, 21), .Cells(12, 21)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 21), .Cells(12, 21)).ColumnWidth = 12
      
      .Range(.Cells(10, 22), .Cells(12, 22)).Merge
      .Range(.Cells(10, 22), .Cells(12, 22)).WrapText = True
      .Range(.Cells(10, 22), .Cells(12, 22)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 22), .Cells(12, 22)).ColumnWidth = 15
      
      .Range(.Cells(10, 23), .Cells(12, 23)).Merge
      .Range(.Cells(10, 23), .Cells(12, 23)).WrapText = True
      .Range(.Cells(10, 23), .Cells(12, 23)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 23), .Cells(12, 23)).ColumnWidth = 8
      
      .Range(.Cells(10, 24), .Cells(12, 24)).Merge
      .Range(.Cells(10, 24), .Cells(12, 24)).WrapText = True
      .Range(.Cells(10, 24), .Cells(12, 24)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 24), .Cells(12, 24)).ColumnWidth = 12
      
      .Range(.Cells(10, 25), .Cells(12, 25)).Merge
      .Range(.Cells(10, 25), .Cells(12, 25)).WrapText = True
      .Range(.Cells(10, 25), .Cells(12, 25)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 25), .Cells(12, 25)).ColumnWidth = 12
      
      .Range(.Cells(10, 26), .Cells(12, 26)).Merge
      .Range(.Cells(10, 26), .Cells(12, 26)).WrapText = True
      .Range(.Cells(10, 26), .Cells(12, 26)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 26), .Cells(12, 26)).ColumnWidth = 10
      
      .Range(.Cells(10, 27), .Cells(12, 27)).Merge
      .Range(.Cells(10, 27), .Cells(12, 27)).WrapText = True
      .Range(.Cells(10, 27), .Cells(12, 27)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 27), .Cells(12, 27)).ColumnWidth = 15
      
      .Range(.Cells(10, 28), .Cells(12, 28)).Merge
      .Range(.Cells(10, 28), .Cells(12, 28)).WrapText = True
      .Range(.Cells(10, 28), .Cells(12, 28)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 28), .Cells(12, 28)).ColumnWidth = 13
      
      .Range(.Cells(10, 29), .Cells(12, 29)).Merge
      .Range(.Cells(10, 29), .Cells(12, 29)).WrapText = True
      .Range(.Cells(10, 29), .Cells(12, 29)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 29), .Cells(12, 29)).ColumnWidth = 13
      
      .Range(.Cells(10, 30), .Cells(12, 30)).Merge
      .Range(.Cells(10, 30), .Cells(12, 30)).WrapText = True
      .Range(.Cells(10, 30), .Cells(12, 30)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 30), .Cells(12, 30)).ColumnWidth = 12
      
      .Range(.Cells(10, 31), .Cells(12, 31)).Merge
      .Range(.Cells(10, 31), .Cells(12, 31)).WrapText = True
      .Range(.Cells(10, 31), .Cells(12, 31)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 31), .Cells(12, 31)).ColumnWidth = 12
      
      .Range(.Cells(10, 32), .Cells(12, 32)).Merge
      .Range(.Cells(10, 32), .Cells(12, 32)).WrapText = True
      .Range(.Cells(10, 32), .Cells(12, 32)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 32), .Cells(12, 32)).ColumnWidth = 12
      
      .Range(.Cells(10, 33), .Cells(12, 33)).Merge
      .Range(.Cells(10, 33), .Cells(12, 33)).WrapText = True
      .Range(.Cells(10, 33), .Cells(12, 33)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(10, 33), .Cells(12, 33)).ColumnWidth = 12
      
      .Cells(10, 1) = "PERIODO"
      .Cells(10, 2) = "TIPO CUO"
      .Cells(10, 3) = "NUMERO CUO"
      .Cells(10, 4) = "FECHA DE EMISION"
      .Cells(10, 5) = "FECHA DE VCTO."
      .Cells(10, 6) = "TIPO (TABLA 10)"
      .Cells(10, 7) = "NRO. SERIE"
      .Cells(10, 8) = "NUMERO INICIAL"
      .Cells(10, 9) = "NUMERO FINAL"
      .Cells(10, 10) = "TIPO (TABLA 2)"
      .Cells(10, 11) = "NUMERO DOC. PROVEEDOR"
      .Cells(10, 12) = "NOMBRE / RAZON SOCIAL"
      .Cells(10, 13) = "BI ADQ. VTAS. GRAV."
      .Cells(10, 14) = "IGV ADQ. VTAS. GRAV."
      .Cells(10, 15) = "BI ADQ. VTAS. GRAV. Y NO GRAV."
      .Cells(10, 16) = "IGV ADQ. VTAS GRAV. Y NO GRAV."
      .Cells(10, 17) = "BI ADQ. VTAS. NO GRAV."
      .Cells(10, 18) = "IGV ADQ. VTAS. NO GRAV."
      .Cells(10, 19) = "ADQ. NO GRAV."
      .Cells(10, 20) = "ISC."
      .Cells(10, 21) = "OTROS"
      .Cells(10, 22) = "IMPORTE TOTAL ADQUISICIONES"
      .Cells(10, 23) = "TIPO DE CAMBIO"
      .Cells(10, 24) = "FECHA EMISION COMPROB. PAGO"
      .Cells(10, 25) = "TIPO COMPROB. PAGO MODIFICA"
      .Cells(10, 26) = "NRO. SERIE COMPROB. MODIFICA"
      .Cells(10, 27) = "CODIGO DEPENDENCIA ADUANERA"
      .Cells(10, 28) = "NRO. COMPROB. MODIFICA"
      .Cells(10, 29) = "NRO. COMPROB. SUJETO NO DOMICILIADO"
      .Cells(10, 30) = "FECHA EMISION CONSTANCIA DEPOSITO"
      .Cells(10, 31) = "NRO CONSTANCIA DEPOSITO O DETRACCION"
      .Cells(10, 32) = "MARCA DEL COMPROB. SUJETO RETENCION"
      .Cells(10, 33) = "ESTADO"
            
      .Range(.Cells(1, 1), .Cells(12, 33)).Font.Bold = True
      .Range(.Cells(10, 1), .Cells(12, 33)).Interior.Color = RGB(146, 208, 80)
      
      .Columns("M").NumberFormat = "#,###,##0.00"
      .Columns("N").NumberFormat = "#,###,##0.00"
      .Columns("O").NumberFormat = "#,###,##0.00"
      .Columns("P").NumberFormat = "#,###,##0.00"
      .Columns("Q").NumberFormat = "#,###,##0.00"
      .Columns("R").NumberFormat = "#,###,##0.00"
      .Columns("S").NumberFormat = "#,###,##0.00"
      .Columns("T").NumberFormat = "#,###,##0.00"
      .Columns("U").NumberFormat = "#,###,##0.00"
      .Columns("V").NumberFormat = "#,###,##0.00"
      .Columns("W").NumberFormat = "###,##0.0000"

      .Range(.Cells(10, 1), .Cells(12, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      Do Until r_int_NroCor > 33
         .Range(.Cells(12, r_int_NroCor), .Cells(12, r_int_NroCor)).VerticalAlignment = xlCenter
         .Range(.Cells(10, r_int_NroCor), .Cells(10, r_int_NroCor)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(12, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(10, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(10, r_int_NroCor), .Cells(10, r_int_NroCor)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(12, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(10, r_int_NroCor), .Cells(12, r_int_NroCor)).Borders(xlEdgeRight).LineStyle = xlContinuous
         r_int_NroCor = r_int_NroCor + 1
      Loop

      r_int_NroCor = 1
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_str_NumRuc = Val(Mid(Trim(g_rst_Princi!CAMPO_11), 1, 11))
         
         If gf_Valida_RUC(r_str_NumRuc, Mid(r_str_NumRuc, 11, 1)) Then
            .Cells(r_int_CntFil, 1) = g_rst_Princi!CAMPO_01
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 1)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 2) = "'" & Format(r_int_NroCor, "000")
            .Range(.Cells(r_int_CntFil, 2), .Cells(r_int_CntFil, 2)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 3) = "M" & Format(r_int_NroCor, "000")
            .Range(.Cells(r_int_CntFil, 3), .Cells(r_int_CntFil, 3)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 4) = "'" & Trim(g_rst_Princi!CAMPO_04)
            .Range(.Cells(r_int_CntFil, 4), .Cells(r_int_CntFil, 4)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 5) = IIf(IsNull(g_rst_Princi!CAMPO_05), "", g_rst_Princi!CAMPO_05)
            
            .Cells(r_int_CntFil, 6) = IIf(IsNull(g_rst_Princi!CAMPO_06), "", g_rst_Princi!CAMPO_06)
            .Range(.Cells(r_int_CntFil, 6), .Cells(r_int_CntFil, 6)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 7) = IIf(IsNull(g_rst_Princi!CAMPO_07), "", g_rst_Princi!CAMPO_07)
            .Range(.Cells(r_int_CntFil, 7), .Cells(r_int_CntFil, 7)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 8) = IIf(IsNull(g_rst_Princi!CAMPO_08), "", g_rst_Princi!CAMPO_08)
            .Range(.Cells(r_int_CntFil, 8), .Cells(r_int_CntFil, 8)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 9) = IIf(IsNull(g_rst_Princi!CAMPO_09), "", g_rst_Princi!CAMPO_09)
            
            .Cells(r_int_CntFil, 10) = IIf(IsNull(g_rst_Princi!CAMPO_10), "", g_rst_Princi!CAMPO_10)
            .Range(.Cells(r_int_CntFil, 10), .Cells(r_int_CntFil, 10)).HorizontalAlignment = xlHAlignCenter
            
            .Cells(r_int_CntFil, 11) = IIf(IsNull(g_rst_Princi!CAMPO_11), "", g_rst_Princi!CAMPO_11)
            .Range(.Cells(r_int_CntFil, 11), .Cells(r_int_CntFil, 11)).HorizontalAlignment = xlHAlignCenter
            
            If InStr(g_rst_Princi!CAMPO_12, "/") - 1 > 0 Then
               .Cells(r_int_CntFil, 12) = Mid(Trim(g_rst_Princi!CAMPO_12), 1, InStr(g_rst_Princi!CAMPO_12, "/") - 1)
            Else
               .Cells(r_int_CntFil, 12) = Mid(Trim(g_rst_Princi!CAMPO_12), 1)
            End If
            .Cells(r_int_CntFil, 13) = IIf(IsNull(g_rst_Princi!CAMPO_13), "", g_rst_Princi!CAMPO_13)
            .Cells(r_int_CntFil, 14) = IIf(IsNull(g_rst_Princi!CAMPO_14), "", g_rst_Princi!CAMPO_14)
            .Cells(r_int_CntFil, 15) = IIf(IsNull(g_rst_Princi!CAMPO_15), "", g_rst_Princi!CAMPO_15)
            .Cells(r_int_CntFil, 16) = IIf(IsNull(g_rst_Princi!CAMPO_16), "", g_rst_Princi!CAMPO_16)
            .Cells(r_int_CntFil, 17) = IIf(IsNull(g_rst_Princi!CAMPO_17), "", g_rst_Princi!CAMPO_17)
            .Cells(r_int_CntFil, 18) = IIf(IsNull(g_rst_Princi!CAMPO_18), "", g_rst_Princi!CAMPO_18)
            .Cells(r_int_CntFil, 19) = IIf(IsNull(g_rst_Princi!CAMPO_19), "", g_rst_Princi!CAMPO_19)
            .Cells(r_int_CntFil, 20) = IIf(IsNull(g_rst_Princi!CAMPO_20), "", g_rst_Princi!CAMPO_20)
            .Cells(r_int_CntFil, 21) = IIf(IsNull(g_rst_Princi!CAMPO_21), "", g_rst_Princi!CAMPO_21)
            .Cells(r_int_CntFil, 22) = IIf(IsNull(g_rst_Princi!CAMPO_22), "", g_rst_Princi!CAMPO_22)
            .Cells(r_int_CntFil, 23) = IIf(IsNull(g_rst_Princi!CAMPO_23), "", g_rst_Princi!CAMPO_23)
            .Cells(r_int_CntFil, 24) = IIf(IsNull(g_rst_Princi!CAMPO_24), "", g_rst_Princi!CAMPO_24)
            .Cells(r_int_CntFil, 25) = IIf(IsNull(g_rst_Princi!CAMPO_25), "", g_rst_Princi!CAMPO_25)
            .Cells(r_int_CntFil, 26) = IIf(IsNull(g_rst_Princi!CAMPO_26), "", g_rst_Princi!CAMPO_26)
            .Cells(r_int_CntFil, 27) = IIf(IsNull(g_rst_Princi!CAMPO_27), "", g_rst_Princi!CAMPO_27)
            .Cells(r_int_CntFil, 28) = IIf(IsNull(g_rst_Princi!CAMPO_28), "", g_rst_Princi!CAMPO_28)
            .Cells(r_int_CntFil, 29) = IIf(IsNull(g_rst_Princi!CAMPO_29), "", g_rst_Princi!CAMPO_29)
            .Cells(r_int_CntFil, 30) = IIf(IsNull(g_rst_Princi!CAMPO_30), "", g_rst_Princi!CAMPO_30)
            .Cells(r_int_CntFil, 31) = IIf(IsNull(g_rst_Princi!CAMPO_31), "", g_rst_Princi!CAMPO_31)
            .Cells(r_int_CntFil, 32) = IIf(IsNull(g_rst_Princi!CAMPO_32), "", g_rst_Princi!CAMPO_32)
            .Cells(r_int_CntFil, 33) = IIf(IsNull(g_rst_Princi!CAMPO_33), "", g_rst_Princi!CAMPO_33)
            
            .Range(.Cells(r_int_CntFil, 29), .Cells(r_int_CntFil, 29)).HorizontalAlignment = xlHAlignCenter
            
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 33)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 33)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 33)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 33)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(r_int_CntFil, 1), .Cells(r_int_CntFil, 33)).Borders(xlInsideVertical).LineStyle = xlContinuous
            
            r_int_NroCor = r_int_NroCor + 1
            r_int_CntFil = r_int_CntFil + 1
         End If
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_CodMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub
