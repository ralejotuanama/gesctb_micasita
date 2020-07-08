VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_GesPer_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   Icon            =   "GesCtb_frm_225.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   4755
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8085
      _Version        =   65536
      _ExtentX        =   14261
      _ExtentY        =   8387
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
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
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
         Begin Threed.SSPanel pnl_Titulo 
            Height          =   300
            Left            =   660
            TabIndex        =   5
            Top             =   150
            Width           =   2955
            _Version        =   65536
            _ExtentX        =   5212
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Gestión Personal"
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
            Picture         =   "GesCtb_frm_225.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   675
         Left            =   60
         TabIndex        =   6
         Top             =   780
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   600
            Left            =   30
            Picture         =   "GesCtb_frm_225.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   600
            Left            =   7320
            Picture         =   "GesCtb_frm_225.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel pnl_Datos 
         Height          =   3180
         Left            =   60
         TabIndex        =   7
         Top             =   1500
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   5609
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
         Begin Threed.SSPanel pnl_CodPla 
            Height          =   315
            Left            =   1590
            TabIndex        =   8
            Top             =   1710
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2822
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin EditLib.fpDoubleSingle ipp_DiaVen 
            Height          =   315
            Left            =   1590
            TabIndex        =   0
            Top             =   2715
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2822
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
            ButtonStyle     =   0
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
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
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
         Begin Threed.SSPanel pnl_Codigo 
            Height          =   315
            Left            =   1590
            TabIndex        =   9
            Top             =   390
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2822
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel lbl_DiaGan 
            Height          =   315
            Left            =   1590
            TabIndex        =   10
            Top             =   2040
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel lbl_DiaGoz 
            Height          =   315
            Left            =   5970
            TabIndex        =   19
            Top             =   2040
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel lbl_DiaSld 
            Height          =   315
            Left            =   1590
            TabIndex        =   21
            Top             =   2370
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel lbl_DiaVig 
            Height          =   315
            Left            =   5970
            TabIndex        =   23
            Top             =   2715
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TipDoc 
            Height          =   315
            Left            =   1590
            TabIndex        =   25
            Top             =   720
            Width           =   5985
            _Version        =   65536
            _ExtentX        =   10557
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_ApeNom 
            Height          =   315
            Left            =   1590
            TabIndex        =   26
            Top             =   1380
            Width           =   5985
            _Version        =   65536
            _ExtentX        =   10557
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   5970
            TabIndex        =   27
            Top             =   1710
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2822
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel pnl_NumDoc 
            Height          =   315
            Left            =   1590
            TabIndex        =   29
            Top             =   1050
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2822
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Documento:"
            Height          =   195
            Left            =   150
            TabIndex        =   30
            Top             =   1140
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ingreso:"
            Height          =   195
            Index           =   1
            Left            =   4530
            TabIndex        =   28
            Top             =   1800
            Width           =   1065
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Vigente (Dias):"
            Height          =   195
            Left            =   4530
            TabIndex        =   24
            Top             =   2820
            Width           =   1035
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Saldo (Dias):"
            Height          =   195
            Left            =   150
            TabIndex        =   22
            Top             =   2490
            Width           =   900
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Gazados (Dias):"
            Height          =   195
            Left            =   4530
            TabIndex        =   20
            Top             =   2160
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   1470
            Width           =   600
         End
         Begin VB.Label lbl_Importe 
            AutoSize        =   -1  'True
            Caption         =   "Vencido (Dias)"
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   2820
            Width           =   1035
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Código Planilla:"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   16
            Top             =   1800
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código Interno:"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   480
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   810
            Width           =   1230
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   90
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   -1410
            TabIndex        =   12
            Top             =   3300
            Width           =   570
         End
         Begin VB.Label lbl_Dia_Sol 
            AutoSize        =   -1  'True
            Caption         =   "Ganados (Dias):"
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   2160
            Width           =   1140
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_GesPer_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Cargar
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0

End Sub

Private Sub fs_Cargar()
Dim r_str_Parame     As String
Dim r_rst_Princi     As ADODB.Recordset
Dim r_int_Ganado     As Integer
Dim r_int_Gozado     As Integer

'   r_str_Parame = ""
'   r_str_Parame = r_str_Parame & " SELECT A.GESPER_CODGES, A.GESPER_TIPDOC, TRIM(B.PARDES_DESCRI) AS TIPO_DOCUMENTO, "
'   r_str_Parame = r_str_Parame & "        A.GESPER_NUMDOC, TRIM(C.MAEPRV_RAZSOC) AS NOMBRE, C.MAEPRV_CODSIC,C.MAEPRV_FECING, "
'   r_str_Parame = r_str_Parame & "        NVL(A.GESPER_DIAGAN,0) AS GESPER_DIAGAN, NVL(A.GESPER_DIAGOZ,0) AS GESPER_DIAGOZ, "
'   r_str_Parame = r_str_Parame & "        NVL(A.GESPER_DIASLD,0) AS GESPER_DIASLD, NVL(A.GESPER_DIAVEN,0) AS GESPER_DIAVEN, "
'   r_str_Parame = r_str_Parame & "        NVL(A.GESPER_DIAVIG,0) AS GESPER_DIAVIG "
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT A.GESPER_CODGES, A.GESPER_TIPDOC, TRIM(B.PARDES_DESCRI) AS TIPO_DOCUMENTO, "
   r_str_Parame = r_str_Parame & "        A.GESPER_NUMDOC, TRIM(C.MAEPRV_RAZSOC) AS NOMBRE, C.MAEPRV_CODSIC,C.MAEPRV_FECING, "
   r_str_Parame = r_str_Parame & "        NVL(A.GESPER_DIAGOZ,0) AS GESPER_DIAGOZ, NVL(A.GESPER_DIAVEN,0) AS GESPER_DIAVEN "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_GESPER A "
   r_str_Parame = r_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 118 AND B.PARDES_CODITE = A.GESPER_TIPDOC "
   r_str_Parame = r_str_Parame & "  INNER JOIN CNTBL_MAEPRV C ON C.MAEPRV_TIPDOC = A.GESPER_TIPDOC AND C.MAEPRV_NUMDOC = A.GESPER_NUMDOC "
   r_str_Parame = r_str_Parame & "  WHERE A.GESPER_TIPDOC = " & moddat_g_int_TipDoc
   r_str_Parame = r_str_Parame & "    AND A.GESPER_NUMDOC = " & moddat_g_str_NumDoc
   r_str_Parame = r_str_Parame & "    AND A.GESPER_TIPTAB = 3 "
   r_str_Parame = r_str_Parame & "    AND A.GESPER_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
      Call frm_Ctb_GesPer_01.fs_SaldoDias(r_rst_Princi!GESPER_TIPDOC, r_rst_Princi!GESPER_NUMDOC, r_int_Ganado, r_int_Gozado)
      
      pnl_Codigo.Caption = CStr(r_rst_Princi!GESPER_CODGES)
      pnl_TipDoc.Caption = Trim(r_rst_Princi!TIPO_DOCUMENTO & "")
      pnl_NumDoc.Caption = Trim(r_rst_Princi!GESPER_NUMDOC & "")
      pnl_ApeNom.Caption = Trim(r_rst_Princi!NOMBRE & "")
      pnl_CodPla.Caption = Trim(r_rst_Princi!MAEPRV_CODSIC & "")
      If Trim(r_rst_Princi!MAEPRV_FECING & "") <> "" Then
         pnl_FecIng.Caption = gf_FormatoFecha(r_rst_Princi!MAEPRV_FECING)
      End If
      lbl_DiaGoz.Caption = r_rst_Princi!GESPER_DIAGOZ & " "
      ipp_DiaVen.Text = r_rst_Princi!GESPER_DIAVEN & " "
      'lbl_DiaGan.Caption = r_rst_Princi!GESPER_DIAGAN & " "
      'lbl_DiaSld.Caption = r_rst_Princi!GESPER_DIASLD & " "
      'lbl_DiaVig.Caption = r_rst_Princi!GESPER_DIAVIG & " "
      
      lbl_DiaGan.Caption = CStr(r_int_Ganado) & " "
      lbl_DiaSld.Caption = CStr(r_int_Ganado - r_rst_Princi!GESPER_DIAGOZ) & " "
      lbl_DiaVig.Caption = CStr((r_int_Ganado - r_rst_Princi!GESPER_DIAGOZ) - r_rst_Princi!GESPER_DIAVEN) & " "
   Else
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT A.MAEPRV_TIPDOC, TRIM(B.PARDES_DESCRI) AS TIPO_DOCUMENTO, "
      r_str_Parame = r_str_Parame & "        A.MAEPRV_NUMDOC, TRIM(A.MAEPRV_RAZSOC) AS NOMBRE, A.MAEPRV_CODSIC, A.MAEPRV_FECING "
      r_str_Parame = r_str_Parame & "   FROM CNTBL_MAEPRV A "
      r_str_Parame = r_str_Parame & "  INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = 118 AND B.PARDES_CODITE = A.MAEPRV_TIPDOC "
      r_str_Parame = r_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & moddat_g_int_TipDoc
      r_str_Parame = r_str_Parame & "    AND A.MAEPRV_NUMDOC = " & moddat_g_str_NumDoc
      r_str_Parame = r_str_Parame & "    AND A.MAEPRV_SITUAC = 1 "
      
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (r_rst_Princi.BOF And r_rst_Princi.EOF) Then
         pnl_Codigo.Caption = ""
         pnl_TipDoc.Caption = Trim(r_rst_Princi!TIPO_DOCUMENTO & "")
         pnl_NumDoc.Caption = Trim(r_rst_Princi!MAEPRV_NUMDOC & "")
         pnl_ApeNom.Caption = Trim(r_rst_Princi!NOMBRE & "")
         pnl_CodPla.Caption = Trim(r_rst_Princi!MAEPRV_CODSIC & "")
         If Trim(r_rst_Princi!MAEPRV_FECING & "") <> "" Then
            pnl_FecIng.Caption = gf_FormatoFecha(r_rst_Princi!MAEPRV_FECING)
         End If
         lbl_DiaGan.Caption = "0 "
         lbl_DiaGoz.Caption = "0 "
         lbl_DiaSld.Caption = "0 "
         ipp_DiaVen.Text = "0 "
         lbl_DiaVig.Caption = "0 "
      End If
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_Parame    As String
Dim r_str_CodGen    As String
Dim r_rst_Genera    As ADODB.Recordset

    If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
       Exit Sub
    End If
 
    'GESTION DE VACACIONES
    'If moddat_g_int_FlgGrb = 1 Then
    If Trim(pnl_Codigo.Caption) = "" Then
       r_str_CodGen = modmip_gf_Genera_CodGen(3, 13)
       moddat_g_int_FlgGrb = 1
    Else
       r_str_CodGen = Trim(pnl_Codigo.Caption)
       moddat_g_int_FlgGrb = 2
    End If

    If Len(Trim(r_str_CodGen)) = 0 Then
       MsgBox "No se genero el código automatico del folio.", vbExclamation, modgen_g_str_NomPlt
       Exit Sub
    End If
   
    r_str_Parame = ""
    r_str_Parame = r_str_Parame & " USP_CNTBL_GESPER ( "
    r_str_Parame = r_str_Parame & CLng(r_str_CodGen) & ", "
    r_str_Parame = r_str_Parame & moddat_g_int_TipDoc & ", "
    r_str_Parame = r_str_Parame & "'" & moddat_g_str_NumDoc & "', "
    r_str_Parame = r_str_Parame & Format(date, "yyyymmdd") & ", "
    r_str_Parame = r_str_Parame & "NULL, " 'TIPO CAMBIO
    r_str_Parame = r_str_Parame & "Null, "
    r_str_Parame = r_str_Parame & "NULL, " 'TIPO MONEDA
    r_str_Parame = r_str_Parame & "Null, "
    r_str_Parame = r_str_Parame & "Null, " 'GESPER_CODBNC
    r_str_Parame = r_str_Parame & "'', " 'GESPER_CTACRR
    r_str_Parame = r_str_Parame & "3 , " 'GESPER_TIPTAB
    r_str_Parame = r_str_Parame & "Null, " 'GESPER_FECHA1
    r_str_Parame = r_str_Parame & "Null, " 'GESPER_FECHA2
    r_str_Parame = r_str_Parame & "'', " 'GESPER_DESCRI
    r_str_Parame = r_str_Parame & CLng(ipp_DiaVen.Text) & ","  'GESPER_DIAVEN
    r_str_Parame = r_str_Parame & CLng(lbl_DiaVig.Caption) & ","  'GESPER_DIAVIG
    r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
    r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
    r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
    r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "') "
    r_str_Parame = r_str_Parame & CStr(moddat_g_int_FlgGrb) & ") " 'as_insupd
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Call frm_Ctb_GesPer_01.fs_BuscarVac
   Screen.MousePointer = 0
   Unload Me
End Sub

Private Sub ipp_DiaVen_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_DiaVen_LostFocus()
   lbl_DiaVig.Caption = CStr(CLng(lbl_DiaSld.Caption) - CLng(ipp_DiaVen.Text)) & " "
End Sub


