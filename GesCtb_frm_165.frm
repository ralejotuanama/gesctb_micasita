VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Pro_ArcRCC_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6945
   ClientLeft      =   5085
   ClientTop       =   3960
   ClientWidth     =   9780
   Icon            =   "GesCtb_frm_165.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9795
      _Version        =   65536
      _ExtentX        =   17277
      _ExtentY        =   12250
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   2175
         Left            =   60
         TabIndex        =   9
         Top             =   3120
         Width           =   9675
         _Version        =   65536
         _ExtentX        =   17066
         _ExtentY        =   3836
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
         Begin VB.FileListBox fil_LisArc 
            Height          =   1845
            Left            =   1590
            TabIndex        =   13
            Top             =   90
            Width           =   3675
         End
         Begin VB.DriveListBox drv_LisUni 
            Height          =   315
            Left            =   5370
            TabIndex        =   11
            Top             =   1800
            Width           =   4245
         End
         Begin VB.DirListBox dir_LisCar 
            Height          =   1665
            Left            =   5370
            TabIndex        =   10
            Top             =   90
            Width           =   4245
         End
         Begin VB.Label Label3 
            Caption         =   "Archivo a cargar:"
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   90
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   825
         Left            =   60
         TabIndex        =   3
         Top             =   1470
         Width           =   9675
         _Version        =   65536
         _ExtentX        =   17066
         _ExtentY        =   1455
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
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   90
            Width           =   8025
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1590
            TabIndex        =   5
            Top             =   420
            Width           =   1125
            _Version        =   196608
            _ExtentX        =   1984
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
            MinValue        =   "2009"
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
            Caption         =   "Mes:"
            Height          =   285
            Left            =   90
            TabIndex        =   7
            Top             =   90
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   285
            Left            =   90
            TabIndex        =   6
            Top             =   420
            Width           =   1305
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   2
         Top             =   780
         Width           =   9675
         _Version        =   65536
         _ExtentX        =   17066
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
            Left            =   60
            Picture         =   "GesCtb_frm_165.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Procesar RCC"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9060
            Picture         =   "GesCtb_frm_165.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   9675
         _Version        =   65536
         _ExtentX        =   17066
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   315
            Left            =   660
            TabIndex        =   17
            Top             =   30
            Width           =   2415
            _Version        =   65536
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Carga de Archivo RCC"
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
         Begin Threed.SSPanel SSPanel9 
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
            Picture         =   "GesCtb_frm_165.frx":0758
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   60
         TabIndex        =   14
         Top             =   6120
         Width           =   9675
         _Version        =   65536
         _ExtentX        =   17066
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
            TabIndex        =   15
            Top             =   360
            Width           =   9555
            _Version        =   65536
            _ExtentX        =   16854
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
         Begin VB.Label lbl_NomPro 
            Caption         =   "Proceso carga informacion Clientes MiCasita"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   90
            Width           =   5505
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   735
         Left            =   60
         TabIndex        =   19
         Top             =   5340
         Width           =   9675
         _Version        =   65536
         _ExtentX        =   17066
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
            TabIndex        =   20
            Top             =   360
            Width           =   9555
            _Version        =   65536
            _ExtentX        =   16854
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
         Begin VB.Label lbl_Erique 
            Caption         =   "Proceso Total:"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   22
            Top             =   90
            Width           =   2535
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   735
         Left            =   60
         TabIndex        =   23
         Top             =   2340
         Width           =   9675
         _Version        =   65536
         _ExtentX        =   17066
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
         Begin VB.ComboBox cmb_TipArc 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   240
            Width           =   3705
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Archivo:"
            Height          =   315
            Left            =   90
            TabIndex        =   24
            Top             =   240
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_ArcRCC_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_lng_TotReg        As Long

Private Type r_Arr_Empresa
   Cod_Empresa          As Integer
   Nom_Empresa          As String
End Type
Dim arr_Empresa()       As r_Arr_Empresa

Private Type r_Arr_Clasif
   Cod_Clasif           As Integer
   Nom_Clasif           As String
End Type
Dim arr_Clasif()        As r_Arr_Clasif

Private Type r_Arr_Credito
   Cod_Credito          As Integer
   Nom_Credito          As String
End Type
Dim arr_Credito()       As r_Arr_Credito

Private Type r_Arr_CargaOpe
   TipNumDoc            As String
   Nom_Cliente          As String
   PerAnoMes            As String
   EmpRep1              As String
   EmpRep2              As String
   ClaDeu               As String
   DiaAtr               As String
   CtaCtb               As String
   IdTipDeu             As String
   TipDeu               As String
   Moneda               As String
   SalDeu1              As String
   SalDeu2              As String
End Type
Dim r_str_TC13()        As r_Arr_CargaOpe

Private Function gf_Buscar_NomEmp(ByVal p_CodEmp As Integer) As String
Dim r_str_Parame  As String

   gf_Buscar_NomEmp = ""
   
   r_str_Parame = "SELECT * FROM CTB_EMPSUP WHERE EMPSUP_CODIGO = " & p_CodEmp & " "
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      gf_Buscar_NomEmp = Trim(g_rst_Listas!EMPSUP_NOMBRE)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function gf_Buscar_TipCla(ByVal p_CodCla As Integer, ByVal p_CodCre As Integer) As String
   gf_Buscar_TipCla = ""
   
   g_str_Parame = "SELECT * FROM CTB_TIPCLA WHERE TIPCLA_TIPCRE = " & p_CodCre & " AND TIPCLA_CODIGO = " & p_CodCla
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      gf_Buscar_TipCla = Trim(g_rst_Listas!TIPCLA_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function gf_Buscar_TipCre_2(ByVal p_CodCre As Integer) As String
   gf_Buscar_TipCre_2 = ""
   
   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE PARDES_CODGRP = '055' AND PARDES_CODITE = " & p_CodCre
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      gf_Buscar_TipCre_2 = Trim(Mid(g_rst_Listas!PARDES_DESCRI, 4, Len(g_rst_Listas!PARDES_DESCRI)))
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Function fs_Busca_Empresas(ByVal p_CodEmp As String) As String
Dim r_int_Contad     As Integer
   
   fs_Busca_Empresas = ""
   For r_int_Contad = 1 To UBound(arr_Empresa)
      If Trim(arr_Empresa(r_int_Contad).Cod_Empresa) = Trim(p_CodEmp) Then
         fs_Busca_Empresas = Trim(arr_Empresa(r_int_Contad).Nom_Empresa)
         Exit For
      End If
   Next r_int_Contad
End Function

Private Function fs_Busca_Clasificacion(ByVal p_CodCla As String) As String
Dim r_int_Contad     As Integer
   
   fs_Busca_Clasificacion = ""
   For r_int_Contad = 1 To UBound(arr_Clasif)
      If Trim(arr_Clasif(r_int_Contad).Cod_Clasif) = Trim(p_CodCla) Then
         fs_Busca_Clasificacion = Trim(arr_Clasif(r_int_Contad).Nom_Clasif)
         Exit For
      End If
   Next r_int_Contad
End Function

Private Function fs_Busca_Creditos(ByVal p_CodCre As String) As String
Dim r_int_Contad     As Integer
   
   fs_Busca_Creditos = "SIN TIPO DE CREDITO"
   For r_int_Contad = 1 To UBound(arr_Credito)
      If Trim(arr_Credito(r_int_Contad).Cod_Credito) = Trim(p_CodCre) Then
         fs_Busca_Creditos = Trim(arr_Credito(r_int_Contad).Nom_Credito)
         Exit For
      End If
   Next r_int_Contad
End Function

Private Sub fs_Carga_Empresas()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CTB_EMPSUP "
   g_str_Parame = g_str_Parame & " ORDER BY EMPSUP_CODIGO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   ReDim arr_Empresa(0)
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ReDim Preserve arr_Empresa(UBound(arr_Empresa) + 1)
         arr_Empresa(UBound(arr_Empresa)).Cod_Empresa = Trim(g_rst_Listas!EMPSUP_CODIGO)
         arr_Empresa(UBound(arr_Empresa)).Nom_Empresa = Trim(g_rst_Listas!EMPSUP_NOMBRE)
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub fs_Carga_Creditos()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CTB_TIPCRE "
   g_str_Parame = g_str_Parame & " ORDER BY TIPCRE_CODIGO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   ReDim arr_Credito(0)
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ReDim Preserve arr_Credito(UBound(arr_Credito) + 1)
         arr_Credito(UBound(arr_Credito)).Cod_Credito = Trim(g_rst_Listas!TIPCRE_CODIGO)
         arr_Credito(UBound(arr_Credito)).Nom_Credito = Trim(g_rst_Listas!TIPCRE_DESCRI)
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub fs_Carga_Clasif()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CTB_TIPCLA "
   g_str_Parame = g_str_Parame & " WHERE TIPCLA_TIPCRE = 13 "
   g_str_Parame = g_str_Parame & "ORDER BY TIPCLA_CODIGO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   ReDim arr_Clasif(0)
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         ReDim Preserve arr_Clasif(UBound(arr_Clasif) + 1)
         arr_Clasif(UBound(arr_Clasif)).Cod_Clasif = Trim(g_rst_Listas!TIPCLA_CODIGO)
         arr_Clasif(UBound(arr_Clasif)).Nom_Clasif = Trim(g_rst_Listas!TIPCLA_DESCRI)
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub

Private Sub fs_CanDat_1(ByVal p_ArcRCC As String)
Dim r_int_NumFil        As Integer
Dim r_str_LineaL        As String
Dim r_int_TipDeu        As Integer
Dim r_str_CarPos        As String

   'Abriendo Archivo RCC
   r_int_NumFil = FreeFile
   Open p_ArcRCC For Input As r_int_NumFil
   
   r_lng_TotReg = 0
   Line Input #r_int_NumFil, r_str_LineaL
   
   Do While Not EOF(r_int_NumFil)
      If Mid(r_str_LineaL, 1, 1) = "1" Then
         Line Input #r_int_NumFil, r_str_LineaL
         DoEvents
         
         r_str_CarPos = Mid(r_str_LineaL, 1, 1)
         
         Do While Not EOF(r_int_NumFil) And r_str_CarPos = Mid(r_str_LineaL, 1, 1)
            r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 2))
         
            If r_int_TipDeu = 13 Then
               'RETIRADO = >(Mid(r_str_LineaL, 19, 4) <> "1418" And Mid(r_str_LineaL, 19, 4) <> "1428" And Mid(r_str_LineaL, 19, 4) <> "1438") And _  Mid(r_str_LineaL, 19, 2) <> "72"
               'ADICIONADO =>(7101, 7103, 7104, 7205, 8104)
               If (Mid(r_str_LineaL, 19, 2) <> "29" And Mid(r_str_LineaL, 19, 2) <> "16" And Mid(r_str_LineaL, 19, 2) <> "84") And _
                  (Mid(r_str_LineaL, 19, 4) <> "1419" And Mid(r_str_LineaL, 19, 4) <> "1429" And Mid(r_str_LineaL, 19, 4) <> "1439") Then
                  
                  If (Mid(r_str_LineaL, 19, 6) = "811302" Or Mid(r_str_LineaL, 19, 6) = "812302" Or Mid(r_str_LineaL, 19, 6) = "813302" Or _
                      Mid(r_str_LineaL, 19, 6) = "811925" Or Mid(r_str_LineaL, 19, 6) = "812925" Or Mid(r_str_LineaL, 19, 6) = "813925" Or _
                      Mid(r_str_LineaL, 19, 6) = "811922" Or Mid(r_str_LineaL, 19, 6) = "812922" Or Mid(r_str_LineaL, 19, 6) = "813922" Or _
                      Mid(r_str_LineaL, 19, 4) = "7111" Or Mid(r_str_LineaL, 19, 4) = "7121" Or _
                      Mid(r_str_LineaL, 19, 4) = "7112" Or Mid(r_str_LineaL, 19, 4) = "7122" Or _
                      Mid(r_str_LineaL, 19, 4) = "7113" Or Mid(r_str_LineaL, 19, 4) = "7123" Or _
                      Mid(r_str_LineaL, 19, 4) = "7114" Or Mid(r_str_LineaL, 19, 4) = "7124" Or _
                      Mid(r_str_LineaL, 19, 4) = "7215" Or Mid(r_str_LineaL, 19, 4) = "7225" Or _
                      Mid(r_str_LineaL, 19, 4) = "8114" Or Mid(r_str_LineaL, 19, 4) = "8124") Or _
                     (Mid(r_str_LineaL, 19, 2) <> "81") Then
                     
                     r_lng_TotReg = r_lng_TotReg + 1
                  End If
               End If
            End If
      
            Line Input #r_int_NumFil, r_str_LineaL
            DoEvents
         Loop

      Else
         'Si es Línea de Detalle
         Line Input #r_int_NumFil, r_str_LineaL
         DoEvents
      End If
   Loop
   
   'Cerrando Archivo RCC
   Close #r_int_NumFil
End Sub

Private Sub fs_CanDat_2(ByVal p_ArcRCC As String)
Dim r_int_NumFil        As Integer
Dim r_str_LineaL        As String
Dim r_int_TipDeu        As Integer
Dim r_str_CarPos        As String

   'Abriendo Archivo RCC
   r_int_NumFil = FreeFile
   Open p_ArcRCC For Input As r_int_NumFil
   
   r_lng_TotReg = 0
   Line Input #r_int_NumFil, r_str_LineaL
   
   Do While Not EOF(r_int_NumFil)
   
      If Mid(r_str_LineaL, 1, 1) = "1" Then
      
         Line Input #r_int_NumFil, r_str_LineaL
         DoEvents
         
         r_str_CarPos = Mid(r_str_LineaL, 1, 1)
         
         Do While Not EOF(r_int_NumFil) And r_str_CarPos = Mid(r_str_LineaL, 1, 1)

            r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 1))
            If r_int_TipDeu = 4 Then
               'RETIRADO => (Mid(r_str_LineaL, 18, 4) <> "1418" And Mid(r_str_LineaL, 18, 4) <> "1428" And Mid(r_str_LineaL, 18, 4) <> "1438") And _ And Mid(r_str_LineaL, 18, 2) <> "72"
               'ADICIONADO => (7101, 7103, 7104, 7205, 8104)
               If (Mid(r_str_LineaL, 18, 2) <> "29" And Mid(r_str_LineaL, 18, 2) <> "16" And Mid(r_str_LineaL, 18, 2) <> "84") And _
                  (Mid(r_str_LineaL, 18, 4) <> "1419" And Mid(r_str_LineaL, 18, 4) <> "1429" And Mid(r_str_LineaL, 18, 4) <> "1439") Then
                  
                  If (Mid(r_str_LineaL, 18, 6) = "811302" Or Mid(r_str_LineaL, 18, 6) = "812302" Or Mid(r_str_LineaL, 18, 6) = "813302" Or _
                      Mid(r_str_LineaL, 18, 6) = "811925" Or Mid(r_str_LineaL, 18, 6) = "812925" Or Mid(r_str_LineaL, 18, 6) = "813925" Or _
                      Mid(r_str_LineaL, 18, 6) = "811922" Or Mid(r_str_LineaL, 18, 6) = "812922" Or Mid(r_str_LineaL, 18, 6) = "813922" Or _
                      Mid(r_str_LineaL, 18, 4) = "7111" Or Mid(r_str_LineaL, 18, 4) = "7121" Or _
                      Mid(r_str_LineaL, 18, 4) = "7112" Or Mid(r_str_LineaL, 18, 4) = "7122" Or _
                      Mid(r_str_LineaL, 18, 4) = "7113" Or Mid(r_str_LineaL, 18, 4) = "7123" Or _
                      Mid(r_str_LineaL, 18, 4) = "7114" Or Mid(r_str_LineaL, 18, 4) = "7124" Or _
                      Mid(r_str_LineaL, 18, 4) = "7215" Or Mid(r_str_LineaL, 18, 4) = "7225" Or _
                      Mid(r_str_LineaL, 18, 4) = "8114" Or Mid(r_str_LineaL, 18, 4) = "8124") Or _
                     (Mid(r_str_LineaL, 18, 2) <> "81") Then
                     
                     r_lng_TotReg = r_lng_TotReg + 1
                     
                  End If
               End If
            End If
      
            Line Input #r_int_NumFil, r_str_LineaL
            DoEvents
         Loop

      Else
         'Si es Línea de Detalle
         Line Input #r_int_NumFil, r_str_LineaL
         DoEvents
      End If
   Loop
   
   'Cerrando Archivo RCC
   Close #r_int_NumFil
End Sub

Private Sub fs_ctbp1004(ByVal p_ArcRCC As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
   Dim r_lng_NumReg        As Long
   Dim r_lng_NumErr        As Long
   Dim r_str_FFnEje        As String
   Dim r_str_HFnEje        As String
   Dim r_str_FecPro        As String
   Dim r_int_NumFil        As Integer
   Dim r_str_LineaL        As String
   Dim r_int_PosSp1        As Integer
   Dim r_int_PosSp2        As Integer
   Dim r_int_PosSp3        As Integer
   Dim r_int_PosSp4        As Integer
   Dim r_int_PosSp5        As Integer
   Dim r_int_PosSp6        As Integer
   Dim r_int_PosSp7        As Integer
   Dim r_int_PosSp8        As Integer
   Dim r_int_PosSp9        As Integer
   Dim r_int_PosS10        As Integer
   Dim r_int_PosS11        As Integer
   Dim r_int_PosS12        As Integer
   Dim r_int_PosS13        As Integer
   Dim r_int_PosS14        As Integer
   Dim r_int_PosS15        As Integer
   Dim r_int_PosS16        As Integer
   Dim r_int_PosS17        As Integer
   Dim r_int_PosS18        As Integer
   Dim r_str_CodSbs        As String
   Dim r_str_FecRep        As String
   Dim r_str_DocTri        As String
   Dim r_str_NumRuc        As String
   Dim r_str_TipDoc        As String
   Dim r_str_NumDoc        As String
   Dim r_str_TipPer        As String
   Dim r_str_Evalua()      As String
   Dim r_str_HipRCC()      As String
   Dim r_int_ConTem        As Integer
   Dim r_int_NumIte        As Integer
   Dim r_dbl_DeuNor        As Double
   Dim r_dbl_DeuCpp        As Double
   Dim r_dbl_DeuDef        As Double
   Dim r_dbl_DeuDud        As Double
   Dim r_dbl_DeuPer        As Double
   Dim r_str_EmpRep        As String
   Dim r_int_MonDeu        As Integer
   Dim r_int_DiaAtr        As Integer
   Dim r_int_TipDeu        As Integer
   Dim r_dbl_SalDeu        As Double
   Dim r_int_ClaDeu        As Integer
   Dim r_str_CtaCtb        As String
   Dim r_int_FlgEnc        As Integer
   Dim r_int_Contad        As Integer
   Dim r_str_CarPos        As String
   Dim r_int_NumEmp        As Integer
   Dim r_str_ApePat        As String
   Dim r_str_ApeMat        As String
   Dim r_str_ApeCas        As String
   Dim r_str_PriNom        As String
   Dim r_str_SegNom        As String
   Dim r_obj_Excel         As Excel.Application
   Dim r_lng_ConVer        As Long
   Dim r_lng_Contad        As Long
   Dim r_lng_ConAux        As Long
   Dim r_lng_ConTem        As Long
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "DETALLE ARCHIVO RCC"
   
   With r_obj_Excel.Sheets(1)
      .Range(.Cells(1, 12), .Cells(1, 13)).Merge
      .Range(.Cells(2, 12), .Cells(2, 13)).Merge
      .Range(.Cells(1, 12), .Cells(2, 13)).Font.Bold = True
      .Cells(1, 12) = "Dpto. de Tecnología e Informática"
      .Cells(2, 12) = "Desarrollo de Sistemas"
           
      .Range(.Cells(4, 1), .Cells(4, 13)).Merge
      .Range(.Cells(4, 1), .Cells(4, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 1)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 1)).Font.Underline = xlUnderlineStyleSingle
      .Cells(4, 1) = "RCC LISTADO DE CREDITOS HIPOTECARIOS - DETALLE"
   
      .Cells(7, 1) = "ITEM"
      .Cells(7, 2) = "DOI CLIENTE"
      .Cells(7, 3) = "NOMBRE CLIENTE"
      .Cells(7, 4) = "PERIODO"
      .Cells(7, 5) = "CODIGO EMPRESA"
      .Cells(7, 6) = "NOMBRE EMPRESA"
      .Cells(7, 7) = "CLASIFICACION"
      .Cells(7, 8) = "DIAS ATRASO"
      .Cells(7, 9) = "CUENTA CONTABLE"
      .Cells(7, 10) = "TIPO DEUDA"
      .Cells(7, 11) = "MONEDA"
      .Cells(7, 12) = "MONTO ($/.)"
      .Cells(7, 13) = "MONTO (US$/.)"
       
      .Range(.Cells(7, 1), .Cells(7, 13)).Font.Bold = True
      .Range(.Cells(7, 1), .Cells(7, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 13)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(7, 1), .Cells(7, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Columns("A").ColumnWidth = 4
      .Columns("B").ColumnWidth = 10
      '.Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 40
      '.Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 7
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 14
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 58
      '.Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 21
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 11
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 16
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("I").NumberFormat = "@"
      .Columns("J").ColumnWidth = 32
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 18
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 15
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 15
      .Columns("M").NumberFormat = "###,###,##0.00"
      
      r_lng_ConVer = 8
      r_lng_NumReg = 0
      r_lng_NumErr = 0
      r_int_NumIte = 0
      ReDim r_str_Evalua(0)
      ReDim r_str_HipRCC(0)
      p_BarPro.FloodPercent = 0
      
      'Fecha de Proceso
      r_str_FecPro = Format(date, "dd/mm/yyyy")
      
      'Abriendo Archivo RCC
      r_int_NumFil = FreeFile
      Open p_ArcRCC For Input As r_int_NumFil
      Line Input #r_int_NumFil, r_str_LineaL
      
      Do While Not EOF(r_int_NumFil)
      
         If Mid(r_str_LineaL, 1, 1) = "1" Then
            r_int_PosSp1 = InStr(1, r_str_LineaL, "|")                           'Código SBS
            r_int_PosSp2 = InStr(r_int_PosSp1 + 1, r_str_LineaL, "|")            'Fecha Reporte
            r_int_PosSp3 = InStr(r_int_PosSp2 + 1, r_str_LineaL, "|")            'Tipo Documento Tributario
            r_int_PosSp4 = InStr(r_int_PosSp3 + 1, r_str_LineaL, "|")            'RUC
            r_int_PosSp5 = InStr(r_int_PosSp4 + 1, r_str_LineaL, "|")            'Tipo Documento de Identidad
            r_int_PosSp6 = InStr(r_int_PosSp5 + 1, r_str_LineaL, "|")            'Número Documento de Identidad
            r_int_PosSp7 = InStr(r_int_PosSp6 + 1, r_str_LineaL, "|")            'Tipo de Persona
            r_int_PosSp8 = InStr(r_int_PosSp7 + 1, r_str_LineaL, "|")            'Tipo de Empresa
            r_int_PosSp9 = InStr(r_int_PosSp8 + 1, r_str_LineaL, "|")            'Cantidad de Empresas
            r_int_PosS10 = InStr(r_int_PosSp9 + 1, r_str_LineaL, "|")            'Deuda Calificación 0
            r_int_PosS11 = InStr(r_int_PosS10 + 1, r_str_LineaL, "|")            'Deuda Calificación 1
            r_int_PosS12 = InStr(r_int_PosS11 + 1, r_str_LineaL, "|")            'Deuda Calificación 2
            r_int_PosS13 = InStr(r_int_PosS12 + 1, r_str_LineaL, "|")            'Deuda Calificación 3
            r_int_PosS14 = InStr(r_int_PosS13 + 1, r_str_LineaL, "|")            'Deuda Calificación 4
            r_int_PosS15 = InStr(r_int_PosS14 + 1, r_str_LineaL, "|")            'Apellido Paterno
            r_int_PosS16 = InStr(r_int_PosS15 + 1, r_str_LineaL, "|")            'Apellido Materno
            r_int_PosS17 = InStr(r_int_PosS16 + 1, r_str_LineaL, "|")            'Apellido Casada
            r_int_PosS18 = InStr(r_int_PosS17 + 1, r_str_LineaL, "|")            'Primer Nombre
            'r_int_PosS19 = InStr(r_int_PosS18 + 1, r_str_LineaL, "|")            'Segundo Nombre
            
            r_str_CodSbs = Mid(r_str_LineaL, 2, r_int_PosSp1 - 2)
            r_str_FecRep = Mid(r_str_LineaL, r_int_PosSp1 + 1, r_int_PosSp2 - 1 - r_int_PosSp1)
            r_str_DocTri = Mid(r_str_LineaL, r_int_PosSp2 + 1, r_int_PosSp3 - 1 - r_int_PosSp2)
            r_str_NumRuc = Mid(r_str_LineaL, r_int_PosSp3 + 1, r_int_PosSp4 - 1 - r_int_PosSp3)
            r_str_TipDoc = Mid(r_str_LineaL, r_int_PosSp4 + 1, r_int_PosSp5 - 1 - r_int_PosSp4)
            r_str_NumDoc = Mid(r_str_LineaL, r_int_PosSp5 + 1, r_int_PosSp6 - 1 - r_int_PosSp5)
            r_str_TipPer = Mid(r_str_LineaL, r_int_PosSp6 + 1, r_int_PosSp7 - 1 - r_int_PosSp6)
            r_str_ApePat = Mid(r_str_LineaL, r_int_PosS14 + 1, r_int_PosS15 - 1 - r_int_PosS14)
            r_str_ApeMat = Mid(r_str_LineaL, r_int_PosS15 + 1, r_int_PosS16 - 1 - r_int_PosS15)
            r_str_ApeCas = Mid(r_str_LineaL, r_int_PosS16 + 1, r_int_PosS17 - 1 - r_int_PosS16)
            r_str_PriNom = Mid(r_str_LineaL, r_int_PosS17 + 1, r_int_PosS18 - 1 - r_int_PosS17)
            r_str_SegNom = Mid(r_str_LineaL, r_int_PosS18 + 1, Len(r_str_LineaL) - r_int_PosS18)
                     
            If Len(Trim(r_str_TipPer)) = 0 Then
               r_str_TipPer = "0"
            End If
            
            r_dbl_DeuNor = 0:    r_dbl_DeuCpp = 0:    r_dbl_DeuDef = 0:    r_dbl_DeuDud = 0:    r_dbl_DeuPer = 0
         
            Line Input #r_int_NumFil, r_str_LineaL
            DoEvents
            
            r_str_CarPos = Mid(r_str_LineaL, 1, 1)
            r_int_NumEmp = 0
            r_int_NumIte = 0
            Erase r_str_Evalua()
            ReDim r_str_Evalua(0)
            
            Do While Not EOF(r_int_NumFil) And r_str_CarPos = Mid(r_str_LineaL, 1, 1)
               r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 1))
                           
               If r_int_TipDeu = 4 Then
                  'RETIRADO => Mid(r_str_LineaL, 18, 4) <> "1418" And Mid(r_str_LineaL, 18, 4) <> "1428" And (Mid(r_str_LineaL, 18, 4) <> "1438") And _ And Mid(r_str_LineaL, 18, 2) <> "72"
                  'ADICIONADO => (7101, 7103, 7104, 7205, 8104)
                                    
                  If (Mid(r_str_LineaL, 18, 2) <> "29" And Mid(r_str_LineaL, 18, 2) <> "16" And Mid(r_str_LineaL, 18, 2) <> "84") And _
                     (Mid(r_str_LineaL, 18, 4) <> "1419" And Mid(r_str_LineaL, 18, 4) <> "1429" And Mid(r_str_LineaL, 18, 4) <> "1439") Then
                     
                     If (Mid(r_str_LineaL, 18, 6) = "811302" Or Mid(r_str_LineaL, 18, 6) = "812302" Or Mid(r_str_LineaL, 18, 6) = "813302" Or _
                         Mid(r_str_LineaL, 18, 6) = "811925" Or Mid(r_str_LineaL, 18, 6) = "812925" Or Mid(r_str_LineaL, 18, 6) = "813925" Or _
                         Mid(r_str_LineaL, 18, 6) = "811922" Or Mid(r_str_LineaL, 18, 6) = "812922" Or Mid(r_str_LineaL, 18, 6) = "813922" Or _
                         Mid(r_str_LineaL, 18, 4) = "7111" Or Mid(r_str_LineaL, 18, 4) = "7121" Or _
                         Mid(r_str_LineaL, 18, 4) = "7112" Or Mid(r_str_LineaL, 18, 4) = "7122" Or _
                         Mid(r_str_LineaL, 18, 4) = "7113" Or Mid(r_str_LineaL, 18, 4) = "7123" Or _
                         Mid(r_str_LineaL, 18, 4) = "7114" Or Mid(r_str_LineaL, 18, 4) = "7124" Or _
                         Mid(r_str_LineaL, 18, 4) = "7215" Or Mid(r_str_LineaL, 18, 4) = "7225" Or _
                         Mid(r_str_LineaL, 18, 4) = "8114" Or Mid(r_str_LineaL, 18, 4) = "8124") Or _
                        (Mid(r_str_LineaL, 18, 2) <> "81") Then
                                    
                        r_int_NumIte = r_int_NumIte + 1
                        If UBound(r_str_Evalua) = 0 Then
                           ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                           r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                           r_int_NumEmp = r_int_NumEmp + 1
                        Else
                           For r_int_ConTem = 1 To UBound(r_str_Evalua) Step 1
                              If r_str_Evalua(r_int_ConTem) = CStr(CLng(Mid(r_str_LineaL, 12, 5))) Then
                                 Exit For
                              End If
                           Next
                            
                           If r_int_ConTem > UBound(r_str_Evalua) Then
                              ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                              r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                              r_int_NumEmp = r_int_NumEmp + 1
                           End If
                        End If
                      
                        r_str_EmpRep = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                        r_int_MonDeu = CInt(Mid(r_str_LineaL, 20, 1))
                        r_int_DiaAtr = CInt(Mid(r_str_LineaL, 32, 4))
                        r_str_CtaCtb = CStr(Mid(r_str_LineaL, 18, 14))
                        r_dbl_SalDeu = CDbl(Mid(r_str_LineaL, 36, 13) & "." & Mid(r_str_LineaL, 49, 2))
                        r_int_ClaDeu = CInt(Mid(r_str_LineaL, 51, 1))
                         
                        Select Case r_int_ClaDeu
                           Case 0:  r_dbl_DeuNor = r_dbl_DeuNor + r_dbl_SalDeu
                           Case 1:  r_dbl_DeuCpp = r_dbl_DeuCpp + r_dbl_SalDeu
                           Case 2:  r_dbl_DeuDef = r_dbl_DeuDef + r_dbl_SalDeu
                           Case 3:  r_dbl_DeuDud = r_dbl_DeuDud + r_dbl_SalDeu
                           Case 4:  r_dbl_DeuPer = r_dbl_DeuPer + r_dbl_SalDeu
                        End Select
                      
                        'Buscando datos de la Garantía en Registro de Hipotecas
                        .Cells(r_lng_ConVer, 1) = r_lng_ConVer - 7
                        .Cells(r_lng_ConVer, 2) = Trim(r_str_TipDoc) & "-" & Trim(r_str_NumDoc)
                        .Cells(r_lng_ConVer, 3) = Trim(r_str_ApePat) & IIf(Len(Trim(r_str_ApeCas)) = 0, " ", Trim(r_str_ApeCas) & " ") & Trim(r_str_ApeMat) & " " & Trim(r_str_PriNom) & " " & Trim(r_str_SegNom)
                        .Cells(r_lng_ConVer, 4) = Trim(p_PerAno) & "-" & Format(Trim(p_PerMes), "00")
                        .Cells(r_lng_ConVer, 5) = Trim(r_str_EmpRep)
                        .Cells(r_lng_ConVer, 6) = gf_Buscar_NomEmp(Trim(r_str_EmpRep))
                        .Cells(r_lng_ConVer, 7) = gf_Buscar_TipCla(Trim(r_int_ClaDeu), 13)
                        .Cells(r_lng_ConVer, 8) = Trim(r_int_DiaAtr)
                        .Cells(r_lng_ConVer, 9) = "" & Trim(r_str_CtaCtb)
                        .Cells(r_lng_ConVer, 10) = gf_Buscar_TipCre_2(Trim(r_int_TipDeu))
                         
                        If r_int_MonDeu = 1 Or r_int_MonDeu = 3 Then
                           .Cells(r_lng_ConVer, 11) = "SOLES"
                           .Cells(r_lng_ConVer, 12) = Format(r_dbl_SalDeu, "###,###,##0.00")
                           .Cells(r_lng_ConVer, 13) = Format(0, "###,###,##0.00")
                        ElseIf r_int_MonDeu = 2 Then
                           .Cells(r_lng_ConVer, 11) = "DOLARES AMERICANOS"
                           .Cells(r_lng_ConVer, 12) = Format(0, "###,###,##0.00")
                           .Cells(r_lng_ConVer, 13) = Format(r_dbl_SalDeu, "###,###,##0.00")
                        End If
                         
                        r_lng_ConVer = r_lng_ConVer + 1
                        r_lng_NumReg = r_lng_NumReg + 1
                        p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
                     End If
                     
                  End If
               End If
         
               Line Input #r_int_NumFil, r_str_LineaL
               DoEvents
            Loop
         
            If r_int_NumEmp > 0 Then
               'Insertar en Arreglo r_str_HipRCC
               ReDim Preserve r_str_HipRCC(UBound(r_str_HipRCC) + IIf(UBound(r_str_HipRCC) = 0, 9, 10))
               r_str_HipRCC(UBound(r_str_HipRCC) - 9) = Trim(r_str_TipDoc) & "-" & Trim(r_str_NumDoc)                'DOI CLIENTE
               r_str_HipRCC(UBound(r_str_HipRCC) - 8) = Trim(r_str_ApePat) & IIf(Len(Trim(r_str_ApeCas)) = 0, " ", Trim(r_str_ApeCas) & " ") & Trim(r_str_ApeMat) & " " & Trim(r_str_PriNom) & " " & Trim(r_str_SegNom)         'NOMBRE CLIENTE
               r_str_HipRCC(UBound(r_str_HipRCC) - 7) = Format(p_PerAno, "0000") & "-" & Format(p_PerMes, "00")      'PERIODO
               r_str_HipRCC(UBound(r_str_HipRCC) - 6) = Trim(r_str_CodSbs)                                           'CODIGO SBS
               r_str_HipRCC(UBound(r_str_HipRCC) - 5) = Trim(r_int_NumEmp)                                           'NUMERO EMPRESAS
               r_str_HipRCC(UBound(r_str_HipRCC) - 4) = Trim(r_dbl_DeuNor)                                           'DEUDA NORMAL
               r_str_HipRCC(UBound(r_str_HipRCC) - 3) = Trim(r_dbl_DeuCpp)                                           'DEUDA CPP
               r_str_HipRCC(UBound(r_str_HipRCC) - 2) = Trim(r_dbl_DeuDef)                                           'DEUDA DEFICIENTE
               r_str_HipRCC(UBound(r_str_HipRCC) - 1) = Trim(r_dbl_DeuDud)                                           'DEUDA DUDOSO
               r_str_HipRCC(UBound(r_str_HipRCC) - 0) = Trim(r_dbl_DeuPer)                                           'DEUDA PERDIDA
               r_lng_TotReg = r_lng_TotReg + 1
            End If
   
         Else
            'Si es Línea de Detalle
            Line Input #r_int_NumFil, r_str_LineaL
            DoEvents
         End If
      Loop
      
      .Range(.Cells(1, 1), .Cells(r_lng_ConVer, 13)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_lng_ConVer, 13)).Font.Size = 8
      .Range(.Cells(1, 12), .Cells(1, 12)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(2, 12), .Cells(2, 12)).HorizontalAlignment = xlHAlignRight
   End With
   
   r_obj_Excel.Sheets(2).Name = "CABECERA ARCHIVO RCC"
   
   With r_obj_Excel.Sheets(2)
      .Range(.Cells(1, 10), .Cells(1, 11)).Merge
      .Range(.Cells(2, 10), .Cells(2, 11)).Merge
      .Range(.Cells(1, 10), .Cells(2, 11)).Font.Bold = True
      .Cells(1, 10) = "Dpto. de Tecnología e Informática"
      .Cells(2, 10) = "Desarrollo de Sistemas"
           
      .Range(.Cells(4, 1), .Cells(4, 11)).Merge
      .Range(.Cells(4, 1), .Cells(4, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 1)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 1)).Font.Underline = xlUnderlineStyleSingle
      .Cells(4, 1) = "RCC LISTADO DE CREDITOS HIPOTECARIOS - CABECERA"
      
      .Cells(7, 1) = "ITEM"
      .Cells(7, 2) = "DOI CLIENTE"
      .Cells(7, 3) = "NOMBRE CLIENTE"
      .Cells(7, 4) = "PERIODO"
      .Cells(7, 5) = "CODIGO SBS"
      .Cells(7, 6) = "NUMERO EMPRESAS"
      .Cells(7, 7) = "DEUDA NORMAL"
      .Cells(7, 8) = "DEUDA CPP"
      .Cells(7, 9) = "DEUDA DEFICIENTE"
      .Cells(7, 10) = "DEUDA DUDOSO"
      .Cells(7, 11) = "DEUDA PERDIDA"
         
      .Range(.Cells(7, 1), .Cells(7, 11)).Font.Bold = True
      .Range(.Cells(7, 1), .Cells(7, 11)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(7, 1), .Cells(7, 11)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(7, 1), .Cells(7, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(7, 1), .Cells(7, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 4
      .Columns("B").ColumnWidth = 10
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 40
      '.Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 7
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 10
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 16
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 15
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 15
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 15
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 15
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 15
      .Columns("K").NumberFormat = "###,###,##0.00"
              
      r_lng_ConVer = 8
      r_lng_Contad = 0
      
      For r_lng_ConAux = 0 To UBound(r_str_HipRCC) Step 10
         .Cells(r_lng_ConVer, 1) = r_lng_ConVer - 7
         For r_lng_ConTem = 2 To 11 Step 1
            .Cells(r_lng_ConVer, r_lng_ConTem) = IIf(r_lng_ConTem < 7, r_str_HipRCC(r_lng_Contad), Format(r_str_HipRCC(r_lng_Contad), "###,###,##0.00"))
            r_lng_Contad = r_lng_Contad + 1
         Next
                                 
         r_lng_ConVer = r_lng_ConVer + 1
         r_lng_NumReg = r_lng_NumReg + 1
         p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
         DoEvents
      Next
      
      .Range(.Cells(1, 1), .Cells(r_lng_ConVer, 13)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_lng_ConVer, 13)).Font.Size = 8
      .Range(.Cells(1, 10), .Cells(1, 10)).HorizontalAlignment = xlHAlignRight
      .Range(.Cells(2, 10), .Cells(2, 10)).HorizontalAlignment = xlHAlignRight
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
   
   'Cerrando Archivo RCC
   Close #r_int_NumFil
End Sub

Private Sub fs_ctbp1003_2(ByVal p_ArcRCC As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
Dim r_int_NumFil        As Integer
Dim r_str_LineaL        As String
Dim r_int_PosSp1        As Integer
Dim r_int_PosSp2        As Integer
Dim r_int_PosSp3        As Integer
Dim r_int_PosSp4        As Integer
Dim r_int_PosSp5        As Integer
Dim r_int_PosSp6        As Integer
Dim r_int_PosSp7        As Integer
Dim r_int_PosSp8        As Integer
Dim r_int_PosSp9        As Integer
Dim r_int_PosS10        As Integer
Dim r_int_PosS11        As Integer
Dim r_int_PosS12        As Integer
Dim r_int_PosS13        As Integer
Dim r_int_PosS14        As Integer
Dim r_int_PosS15        As Integer
Dim r_int_PosS16        As Integer
Dim r_int_PosS17        As Integer
Dim r_int_PosS18        As Integer
Dim r_str_CodSbs        As String
Dim r_str_FecRep        As String
Dim r_str_DocTri        As String
Dim r_str_NumRuc        As String
Dim r_str_TipDoc        As String
Dim r_str_NumDoc        As String
Dim r_str_TipPer        As String
Dim r_str_TipEmp        As String
Dim r_str_Evalua()      As String
Dim r_int_ConTem        As Integer
Dim r_int_NumIte        As Long
Dim r_lng_Contad        As Long
Dim r_dbl_DeuNor        As Double
Dim r_dbl_DeuCpp        As Double
Dim r_dbl_DeuDef        As Double
Dim r_dbl_DeuDud        As Double
Dim r_dbl_DeuPer        As Double
Dim r_str_EmpRep        As String
Dim r_int_MonDeu        As Integer
Dim r_int_DiaAtr        As Integer
Dim r_int_TipDeu        As Integer
Dim r_dbl_SalDeu        As Double
Dim r_int_ClaDeu        As Integer
Dim r_str_CtaCtb        As String
Dim r_str_CarPos        As String
Dim r_int_NumEmp        As Long
Dim r_str_ApePat        As String
Dim r_str_ApeMat        As String
Dim r_str_ApeCas        As String
Dim r_str_PriNom        As String
Dim r_str_SegNom        As String
Dim TipDeu              As Boolean
   
   r_lng_Contad = 1
   r_int_NumIte = 8
   ReDim r_str_TC13(0)
   ReDim r_str_Evalua(0)
   p_BarPro.FloodPercent = 0
   
   'Abriendo Archivo RCC
   r_int_NumFil = FreeFile
   Open p_ArcRCC For Input As r_int_NumFil
   
   Line Input #r_int_NumFil, r_str_LineaL
   Do While Not EOF(r_int_NumFil)
   
      If Mid(r_str_LineaL, 1, 1) = "1" Then
         TipDeu = False
         r_int_PosSp1 = InStr(1, r_str_LineaL, "|")                           'Código SBS
         r_int_PosSp2 = InStr(r_int_PosSp1 + 1, r_str_LineaL, "|")            'Fecha Reporte
         r_int_PosSp3 = InStr(r_int_PosSp2 + 1, r_str_LineaL, "|")            'Tipo Documento Tributario
         r_int_PosSp4 = InStr(r_int_PosSp3 + 1, r_str_LineaL, "|")            'RUC
         r_int_PosSp5 = InStr(r_int_PosSp4 + 1, r_str_LineaL, "|")            'Tipo Documento de Identidad
         r_int_PosSp6 = InStr(r_int_PosSp5 + 1, r_str_LineaL, "|")            'Número Documento de Identidad
         r_int_PosSp7 = InStr(r_int_PosSp6 + 1, r_str_LineaL, "|")            'Tipo de Persona
         r_int_PosSp8 = InStr(r_int_PosSp7 + 1, r_str_LineaL, "|")            'Tipo de Empresa
         r_str_TipPer = Mid(r_str_LineaL, r_int_PosSp6 + 1, r_int_PosSp7 - 1 - r_int_PosSp6)
         
         'Condicion diferente a Empresas
         If Trim(r_str_TipPer) = 1 Or Trim(r_str_TipPer) = 2 Or Trim(r_str_TipPer) = 3 Then
            r_int_PosSp9 = InStr(r_int_PosSp8 + 1, r_str_LineaL, "|")            'Cantidad de Empresas
            r_int_PosS10 = InStr(r_int_PosSp9 + 1, r_str_LineaL, "|")            'Deuda Calificación 0
            r_int_PosS11 = InStr(r_int_PosS10 + 1, r_str_LineaL, "|")            'Deuda Calificación 1
            r_int_PosS12 = InStr(r_int_PosS11 + 1, r_str_LineaL, "|")            'Deuda Calificación 2
            r_int_PosS13 = InStr(r_int_PosS12 + 1, r_str_LineaL, "|")            'Deuda Calificación 3
            r_int_PosS14 = InStr(r_int_PosS13 + 1, r_str_LineaL, "|")            'Deuda Calificación 4
            r_int_PosS15 = InStr(r_int_PosS14 + 1, r_str_LineaL, "|")            'Apellido Paterno
            r_int_PosS16 = InStr(r_int_PosS15 + 1, r_str_LineaL, "|")            'Apellido Materno
            r_int_PosS17 = InStr(r_int_PosS16 + 1, r_str_LineaL, "|")            'Apellido Casada
            r_int_PosS18 = InStr(r_int_PosS17 + 1, r_str_LineaL, "|")            'Primer Nombre
            
            r_str_CodSbs = Mid(r_str_LineaL, 2, r_int_PosSp1 - 2)
            r_str_FecRep = Mid(r_str_LineaL, r_int_PosSp1 + 1, r_int_PosSp2 - 1 - r_int_PosSp1)
            r_str_DocTri = Mid(r_str_LineaL, r_int_PosSp2 + 1, r_int_PosSp3 - 1 - r_int_PosSp2)
            r_str_NumRuc = Mid(r_str_LineaL, r_int_PosSp3 + 1, r_int_PosSp4 - 1 - r_int_PosSp3)
            r_str_TipDoc = Mid(r_str_LineaL, r_int_PosSp4 + 1, r_int_PosSp5 - 1 - r_int_PosSp4)
            r_str_NumDoc = Mid(r_str_LineaL, r_int_PosSp5 + 1, r_int_PosSp6 - 1 - r_int_PosSp5)
            r_str_TipEmp = Mid(r_str_LineaL, r_int_PosSp7 + 1, r_int_PosSp8 - 1 - r_int_PosSp7)
            r_str_ApePat = Mid(r_str_LineaL, r_int_PosS14 + 1, r_int_PosS15 - 1 - r_int_PosS14)
            r_str_ApeMat = Mid(r_str_LineaL, r_int_PosS15 + 1, r_int_PosS16 - 1 - r_int_PosS15)
            r_str_ApeCas = Mid(r_str_LineaL, r_int_PosS16 + 1, r_int_PosS17 - 1 - r_int_PosS16)
            r_str_PriNom = Mid(r_str_LineaL, r_int_PosS17 + 1, r_int_PosS18 - 1 - r_int_PosS17)
            r_str_SegNom = Mid(r_str_LineaL, r_int_PosS18 + 1, Len(r_str_LineaL) - r_int_PosS18)
            
            r_dbl_DeuNor = 0:    r_dbl_DeuCpp = 0:    r_dbl_DeuDef = 0:    r_dbl_DeuDud = 0:    r_dbl_DeuPer = 0
            Line Input #r_int_NumFil, r_str_LineaL
            DoEvents
            
            r_str_CarPos = Mid(r_str_LineaL, 1, 1)
            r_int_NumEmp = 0
            
            Erase r_str_Evalua()
            ReDim r_str_Evalua(0)
            Erase r_str_TC13()
            ReDim r_str_TC13(0)
            
            Do While Not EOF(r_int_NumFil) And r_str_CarPos = Mid(r_str_LineaL, 1, 1)
               If Mid(r_str_LineaL, 19, 8) = "14110302" Or Mid(r_str_LineaL, 19, 8) = "14210302" Then
                  r_int_TipDeu = 9                                   'Tarjeta de Crédito
               Else
                  r_int_TipDeu = CInt(Mid(r_str_LineaL, 17, 2))
               End If
               'RETIRADO => (Mid(r_str_LineaL, 19, 4) <> "1418" And Mid(r_str_LineaL, 19, 4) <> "1428" And Mid(r_str_LineaL, 19, 4) <> "1438") And _
               'ADICIONADO => (7101, 7103, 7104, 7205, 8104) And Mid(r_str_LineaL, 19, 2) <> "72"
               If (Mid(r_str_LineaL, 19, 2) <> "29" And Mid(r_str_LineaL, 19, 2) <> "16" And Mid(r_str_LineaL, 19, 2) <> "84") And _
                  (Mid(r_str_LineaL, 19, 4) <> "1419" And Mid(r_str_LineaL, 19, 4) <> "1429" And Mid(r_str_LineaL, 19, 4) <> "1439") Then
                  
                  If (Mid(r_str_LineaL, 19, 6) = "811302" Or Mid(r_str_LineaL, 19, 6) = "812302" Or Mid(r_str_LineaL, 19, 6) = "813302" Or _
                      Mid(r_str_LineaL, 19, 6) = "811925" Or Mid(r_str_LineaL, 19, 6) = "812925" Or Mid(r_str_LineaL, 19, 6) = "813925" Or _
                      Mid(r_str_LineaL, 19, 6) = "811922" Or Mid(r_str_LineaL, 19, 6) = "812922" Or Mid(r_str_LineaL, 19, 6) = "813922" Or _
                      Mid(r_str_LineaL, 19, 4) = "7111" Or Mid(r_str_LineaL, 19, 4) = "7121" Or _
                      Mid(r_str_LineaL, 19, 4) = "7112" Or Mid(r_str_LineaL, 19, 4) = "7122" Or _
                      Mid(r_str_LineaL, 19, 4) = "7113" Or Mid(r_str_LineaL, 19, 4) = "7123" Or _
                      Mid(r_str_LineaL, 19, 4) = "7114" Or Mid(r_str_LineaL, 19, 4) = "7124" Or _
                      Mid(r_str_LineaL, 19, 4) = "7215" Or Mid(r_str_LineaL, 19, 4) = "7225" Or _
                      Mid(r_str_LineaL, 19, 4) = "8114" Or Mid(r_str_LineaL, 19, 4) = "8124") Or _
                     (Mid(r_str_LineaL, 19, 2) <> "81") Then
                     
                     If UBound(r_str_Evalua) = 0 Then
                        ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                        r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                        r_int_NumEmp = r_int_NumEmp + 1
                     Else
                        For r_int_ConTem = 1 To UBound(r_str_Evalua) Step 1
                           If r_str_Evalua(r_int_ConTem) = CStr(CLng(Mid(r_str_LineaL, 12, 5))) Then
                              Exit For
                           End If
                        Next
                        If r_int_ConTem > UBound(r_str_Evalua) Then
                           ReDim Preserve r_str_Evalua(UBound(r_str_Evalua) + 1)
                           r_str_Evalua(UBound(r_str_Evalua)) = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                           r_int_NumEmp = r_int_NumEmp + 1
                        End If
                     End If
                     
                     r_str_EmpRep = CStr(CLng(Mid(r_str_LineaL, 12, 5)))
                     r_int_MonDeu = CInt(Mid(r_str_LineaL, 21, 1))
                     r_int_DiaAtr = CInt(Mid(r_str_LineaL, 33, 4))
                     r_str_CtaCtb = CStr(Mid(r_str_LineaL, 19, 14))
                     r_dbl_SalDeu = CDbl(Mid(r_str_LineaL, 37, 16) & "." & Mid(r_str_LineaL, 53, 2))
                     r_int_ClaDeu = CInt(Mid(r_str_LineaL, 55, 1))
                     
                     Select Case r_int_ClaDeu
                        Case 0:  r_dbl_DeuNor = r_dbl_DeuNor + r_dbl_SalDeu
                        Case 1:  r_dbl_DeuCpp = r_dbl_DeuCpp + r_dbl_SalDeu
                        Case 2:  r_dbl_DeuDef = r_dbl_DeuDef + r_dbl_SalDeu
                        Case 3:  r_dbl_DeuDud = r_dbl_DeuDud + r_dbl_SalDeu
                        Case 4:  r_dbl_DeuPer = r_dbl_DeuPer + r_dbl_SalDeu
                     End Select
                     
                     If r_int_TipDeu = 13 Then
                        TipDeu = True
                     End If
                     
                     ReDim Preserve r_str_TC13(UBound(r_str_TC13) + 1)
                     r_str_TC13(UBound(r_str_TC13)).TipNumDoc = Trim(r_str_TipDoc) & "-" & Trim(r_str_NumDoc)
                     r_str_TC13(UBound(r_str_TC13)).Nom_Cliente = Trim(r_str_ApePat) & IIf(Len(Trim(r_str_ApeCas)) = 0, " ", Trim(r_str_ApeCas) & " ") & Trim(r_str_ApeMat) & " " & Trim(r_str_PriNom) & " " & Trim(r_str_SegNom)
                     r_str_TC13(UBound(r_str_TC13)).PerAnoMes = Trim(p_PerAno) & "-" & Format(Trim(p_PerMes), "00")
                     r_str_TC13(UBound(r_str_TC13)).EmpRep1 = Trim(r_str_EmpRep)
                     r_str_TC13(UBound(r_str_TC13)).EmpRep2 = fs_Busca_Empresas(Trim(r_str_EmpRep))
                     r_str_TC13(UBound(r_str_TC13)).ClaDeu = fs_Busca_Clasificacion(Trim(r_int_ClaDeu))
                     r_str_TC13(UBound(r_str_TC13)).DiaAtr = Trim(r_int_DiaAtr)
                     r_str_TC13(UBound(r_str_TC13)).CtaCtb = " " & Trim(r_str_CtaCtb)
                     r_str_TC13(UBound(r_str_TC13)).IdTipDeu = Trim(r_int_TipDeu)
                     r_str_TC13(UBound(r_str_TC13)).TipDeu = fs_Busca_Creditos(Trim(r_int_TipDeu))
                     If r_int_MonDeu = 1 Or r_int_MonDeu = 3 Then
                        r_str_TC13(UBound(r_str_TC13)).Moneda = "SOLES"
                        r_str_TC13(UBound(r_str_TC13)).SalDeu1 = Format(r_dbl_SalDeu, "###,###,##0.00")
                        r_str_TC13(UBound(r_str_TC13)).SalDeu2 = Format(0, "###,###,##0.00")
                     ElseIf r_int_MonDeu = 2 Then
                        r_str_TC13(UBound(r_str_TC13)).Moneda = "DOLARES AMERICANOS"
                        r_str_TC13(UBound(r_str_TC13)).SalDeu1 = Format(0, "###,###,##0.00")
                        r_str_TC13(UBound(r_str_TC13)).SalDeu2 = Format(r_dbl_SalDeu, "###,###,##0.00")
                     End If
                  End If
               End If
               
               Line Input #r_int_NumFil, r_str_LineaL
               DoEvents
            Loop
               
            If TipDeu = True Then
               Dim i As Integer
               For i = 1 To UBound(r_str_TC13) Step 1
                  'Insertando Registro Detalle
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & "INSERT INTO RCC_HIPDET("
                  g_str_Parame = g_str_Parame & "HIPDET_PERANO, "
                  g_str_Parame = g_str_Parame & "HIPDET_PERMES, "
                  g_str_Parame = g_str_Parame & "HIPDET_CODEMP, "
                  g_str_Parame = g_str_Parame & "HIPDET_TIPDEU, "
                  g_str_Parame = g_str_Parame & "HIPDET_CTACBL, "
                  g_str_Parame = g_str_Parame & "HIPDET_NUMITE, "
                  g_str_Parame = g_str_Parame & "HIPDET_TIPPER, "
                  g_str_Parame = g_str_Parame & "HIPDET_TIPDOC, "
                  g_str_Parame = g_str_Parame & "HIPDET_DOCIDE, "
                  g_str_Parame = g_str_Parame & "HIPDET_NOMCLI, "
                  g_str_Parame = g_str_Parame & "HIPDET_NOMEMP, "
                  g_str_Parame = g_str_Parame & "HIPDET_CLASIF, "
                  g_str_Parame = g_str_Parame & "HIPDET_DIAATR, "
                  g_str_Parame = g_str_Parame & "HIPDET_TIPMON, "
                  g_str_Parame = g_str_Parame & "HIPDET_MTOSOL, "
                  g_str_Parame = g_str_Parame & "HIPDET_MTODOL) "
                  g_str_Parame = g_str_Parame & "VALUES ("
                  g_str_Parame = g_str_Parame & "'" & Format(p_PerAno, "0000") & "', "
                  g_str_Parame = g_str_Parame & "'" & Format(p_PerMes, "00") & "', "
                  g_str_Parame = g_str_Parame & Trim(r_str_TC13(i).EmpRep1) & ", "
                  g_str_Parame = g_str_Parame & "'" & Trim(r_str_TC13(i).TipDeu) & "', "
                  g_str_Parame = g_str_Parame & "'" & Trim(r_str_TC13(i).CtaCtb) & "', "
                  g_str_Parame = g_str_Parame & CStr(r_lng_Contad) & ", "
                  g_str_Parame = g_str_Parame & "'" & Trim(r_str_TipPer) & "', "
                  If CInt(r_str_TipPer) = 1 Or CInt(r_str_TipPer) = 3 Then
                     g_str_Parame = g_str_Parame & "'" & Trim(r_str_TipDoc) & "', "
                     g_str_Parame = g_str_Parame & "'" & Trim(r_str_NumDoc) & "', "
                  ElseIf CInt(r_str_TipPer) = 2 Then
                     g_str_Parame = g_str_Parame & "'" & Trim(r_str_DocTri) & "', "
                     g_str_Parame = g_str_Parame & "'" & Trim(r_str_NumRuc) & "', "
                  End If
                  g_str_Parame = g_str_Parame & "'" & Trim(Replace(Trim(r_str_ApePat) & IIf(Len(Trim(r_str_ApeCas)) = 0, " ", Trim(r_str_ApeCas) & " ") & Trim(r_str_ApeMat) & " " & Trim(r_str_PriNom) & " " & Trim(r_str_SegNom), "'", "''")) & "', "
                  g_str_Parame = g_str_Parame & "'" & Trim(r_str_TC13(i).EmpRep2) & "', "
                  g_str_Parame = g_str_Parame & "'" & Trim(r_str_TC13(i).ClaDeu) & "', "
                  g_str_Parame = g_str_Parame & Trim(r_str_TC13(i).DiaAtr) & ", "
                  g_str_Parame = g_str_Parame & "'" & Trim(r_str_TC13(i).Moneda) & "', "
                  g_str_Parame = g_str_Parame & Format(Trim(r_str_TC13(i).SalDeu1), "########0.00") & ", "
                  g_str_Parame = g_str_Parame & Format(Trim(r_str_TC13(i).SalDeu2), "########0.00") & ")"
                                           
                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                     Exit Sub
                  End If
                  r_lng_Contad = r_lng_Contad + 1
               Next
               
               '----------------------------------------------------------------------------
               'Insertando Registro Cabecera
               g_str_Parame = ""
               g_str_Parame = g_str_Parame & "INSERT INTO RCC_HIPCAB("
               g_str_Parame = g_str_Parame & "HIPCAB_PERANO, "
               g_str_Parame = g_str_Parame & "HIPCAB_PERMES, "
               g_str_Parame = g_str_Parame & "HIPCAB_CODSBS, "
               g_str_Parame = g_str_Parame & "HIPCAB_TIPDOC, "
               g_str_Parame = g_str_Parame & "HIPCAB_DOCIDE, "
               g_str_Parame = g_str_Parame & "HIPCAB_TIPPER, "
               g_str_Parame = g_str_Parame & "HIPCAB_NOMCLI, "
               g_str_Parame = g_str_Parame & "HIPCAB_NUMEMP, "
               g_str_Parame = g_str_Parame & "HIPCAB_DEUNOR, "
               g_str_Parame = g_str_Parame & "HIPCAB_DEUCPP, "
               g_str_Parame = g_str_Parame & "HIPCAB_DEUDEF, "
               g_str_Parame = g_str_Parame & "HIPCAB_DEUDUD, "
               g_str_Parame = g_str_Parame & "HIPCAB_DEUPER) "
               g_str_Parame = g_str_Parame & "VALUES ("
               g_str_Parame = g_str_Parame & "'" & Format(p_PerAno, "0000") & "', "
               g_str_Parame = g_str_Parame & "'" & Format(p_PerMes, "00") & "', "
               g_str_Parame = g_str_Parame & "'" & Trim(r_str_CodSbs) & "', "
               g_str_Parame = g_str_Parame & "'" & Trim(r_str_TipDoc) & "', "
               g_str_Parame = g_str_Parame & "'" & Trim(r_str_NumDoc) & "', "
               g_str_Parame = g_str_Parame & "'" & Trim(r_str_TipPer) & "', "
               g_str_Parame = g_str_Parame & "'" & Trim(Replace(Trim(r_str_ApePat) & IIf(Len(Trim(r_str_ApeCas)) = 0, " ", Trim(r_str_ApeCas) & " ") & Trim(r_str_ApeMat) & " " & Trim(r_str_PriNom) & " " & Trim(r_str_SegNom), "'", "''")) & "', "
               g_str_Parame = g_str_Parame & "'" & Trim(r_int_NumEmp) & "', "
               g_str_Parame = g_str_Parame & Format(Trim(r_dbl_DeuNor), "########0.00") & ", "
               g_str_Parame = g_str_Parame & Format(Trim(r_dbl_DeuCpp), "########0.00") & ", "
               g_str_Parame = g_str_Parame & Format(Trim(r_dbl_DeuDef), "########0.00") & ", "
               g_str_Parame = g_str_Parame & Format(Trim(r_dbl_DeuDud), "########0.00") & ", "
               g_str_Parame = g_str_Parame & Format(Trim(r_dbl_DeuPer), "########0.00") & ")"
                                        
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                  Exit Sub
               End If
               '----------------------------------------------------------------------------
               
               r_int_NumIte = r_int_NumIte + 1
               p_BarPro.FloodPercent = CDbl(Format(r_int_NumIte / r_lng_TotReg * 100, "##0.00"))
            End If
            
         Else
            Line Input #r_int_NumFil, r_str_LineaL
            DoEvents
         End If
         
      Else
         'Si es Línea de Detalle
         Line Input #r_int_NumFil, r_str_LineaL
         DoEvents
      End If
      
      TipDeu = False
   Loop
   
   'Cerrando Archivo RCC
   Close #r_int_NumFil
End Sub

Private Sub cmb_PerMes_Click()
   Call gs_SetFocus(ipp_PerAno)
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PerMes_Click
   End If
End Sub

Private Sub cmb_TipArc_Click()
   If CInt(cmb_TipArc.ItemData(cmb_TipArc.ListIndex)) = 0 Then
      fil_LisArc.Pattern = "rcc*.ope"
   ElseIf CInt(cmb_TipArc.ItemData(cmb_TipArc.ListIndex)) = 1 Then
      fil_LisArc.Pattern = "*.xls"
   End If
End Sub

Private Sub cmd_Proces_Click()
Dim r_lng_TotErr     As Long
Dim r_lng_Contad     As Long
Dim r_str_Fecha1     As String
Dim r_str_Fecha2     As String
Dim r_str_Fecha3     As String

   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If cmb_TipArc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Archivo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipArc)
      Exit Sub
   End If
   If Len(Trim(fil_LisArc.FileName & "")) = 0 Then
      MsgBox "Debe seleccionar el Archivo a cargar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de cargar la información del RCC?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   pnl_BarTot.FloodPercent = 0
   cmd_Proces.Enabled = False
   
   If CInt(cmb_TipArc.ItemData(cmb_TipArc.ListIndex)) = 0 Then
      'Validar que no se haya cargado información de este Período
      r_str_Fecha1 = Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss")
      lbl_NomPro.Caption = "Verificando si existe data para el periodo (clientes miCasita)...": DoEvents
      g_str_Parame = "SELECT NVL(COUNT(*),0) AS TOTREG FROM CLI_RCCCAB WHERE RCCCAB_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND RCCCAB_PERANO = " & CStr(ipp_PerAno.Text) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      r_lng_Contad = g_rst_Princi!TOTREG
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      '***** PARTE 1/3
      If r_lng_Contad > 0 Then
         If MsgBox("La información de los clientes micasita del RCC para este Período ya ha sido cargada. ¿Desea volver a cargar la información?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
            'Borra
            lbl_NomPro.Caption = "Eliminando información Clientes miCasita...": DoEvents
            Call modprc_ctbp1004("CTBP1004", Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_TotErr, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
            
            'Carga
            lbl_NomPro.Caption = "Proceso carga información Clientes miCasita...": DoEvents
            Call modprc_ctbp1003("CTBP1003", Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_TotErr, fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
         End If
      Else
         'Carga
         lbl_NomPro.Caption = "Proceso carga información Clientes miCasita...": DoEvents
         Call modprc_ctbp1003("CTBP1003", Format(date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_TotErr, fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
      End If
      
      pnl_BarTot.FloodPercent = 35
      
      '***** PARTE 2/3
      'Validar que no se haya cargado información de este Período
      r_str_Fecha2 = Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss")
      lbl_NomPro.Caption = "Verificando si existe data para el periodo (No clientes miCasita)...": DoEvents
      g_str_Parame = "SELECT NVL(COUNT(*),0) AS TOTREG FROM RCC_HIPCAB WHERE HIPCAB_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND HIPCAB_PERANO = " & CStr(ipp_PerAno.Text) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      r_lng_Contad = g_rst_Princi!TOTREG
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      If r_lng_Contad > 0 Then
         If MsgBox("La información de los no clientes de micasita del RCC para este Período ya ha sido cargada. ¿Desea volver a cargar la información?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
            lbl_NomPro.Caption = "Eliminando informacion no clientes miCasita...": DoEvents
            
            'Borra Cabecera
            pnl_BarPro.FloodPercent = 0
            modprc_g_str_CadEje = "DELETE FROM RCC_HIPCAB WHERE "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCAB_PERANO = " & CStr(ipp_PerAno.Text) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCAB_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
            
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 2) Then
               Exit Sub
            End If
            
            'Borra Detalle
            pnl_BarPro.FloodPercent = 50
            modprc_g_str_CadEje = "DELETE FROM RCC_HIPDET WHERE "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPDET_PERANO = " & CStr(ipp_PerAno.Text) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPDET_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 2) Then
               Exit Sub
            End If
               
            pnl_BarPro.FloodPercent = 100
            
            'Carga
            lbl_NomPro.Caption = "Proceso carga informacion No Clientes MiCasita con Creditos Hipotecarios...": DoEvents
            If cmb_PerMes.ItemData(cmb_PerMes.ListIndex) <= 6 And CInt(ipp_PerAno.Text) <= 2010 Then
               Call fs_CanDat_2(fil_LisArc.Path & "\" & fil_LisArc.FileName)
               Call fs_ctbp1004(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
            Else
               Call fs_CanDat_1(fil_LisArc.Path & "\" & fil_LisArc.FileName)
               Call fs_ctbp1003_2(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
            End If
         End If
      Else
         'Carga
         lbl_NomPro.Caption = "Proceso carga información No Clientes MiCasita con Creditos Hipotecarios...": DoEvents
         If cmb_PerMes.ItemData(cmb_PerMes.ListIndex) <= 6 And CInt(ipp_PerAno.Text) <= 2010 Then
            Call fs_CanDat_2(fil_LisArc.Path & "\" & fil_LisArc.FileName)
            Call fs_ctbp1004(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
         Else
            Call fs_CanDat_1(fil_LisArc.Path & "\" & fil_LisArc.FileName)
            Call fs_ctbp1003_2(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
         End If
      End If
      
      '***** PARTE 3/3
      pnl_BarTot.FloodPercent = 70
      
      'Validar que no se haya cargado información (Saldos, Vencidos, M90, Judiciales, Castigados, HT y HP) de este Período
      r_str_Fecha2 = Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss")
      lbl_NomPro.Caption = "Verificando si existe data para el periodo (Según tipo de datos)...": DoEvents
      g_str_Parame = "SELECT NVL(COUNT(*),0) AS TOTREG FROM RCC_CONMEN WHERE CONMEN_PERANO = " & CStr(ipp_PerAno.Text) & " AND CONMEN_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & "  "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox g_str_Parame
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      r_lng_Contad = g_rst_Princi!TOTREG
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      If r_lng_Contad > 0 Then
         If MsgBox("La información según Tipo de Datos del RCC para este Período ya ha sido cargada. ¿Desea volver a cargar la información?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
            lbl_NomPro.Caption = "Eliminando informacion según Tipo de Dato...": DoEvents
               
            'Borra
            pnl_BarPro.FloodPercent = 50
            modprc_g_str_CadEje = "DELETE FROM RCC_CONMEN WHERE "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "CONMEN_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "CONMEN_PERANO = " & CStr(ipp_PerAno.Text) & " "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 2) Then
               MsgBox g_str_Parame
               Exit Sub
            End If
            
            'Carga
            pnl_BarPro.FloodPercent = 100
            lbl_NomPro.Caption = "Proceso carga información según Tipo de Datos con Creditos Hipotecarios...": DoEvents
            Call fs_CargaTipDat(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
         End If
      Else
         'Carga
         lbl_NomPro.Caption = "Proceso carga información según Tipo de Datos con Creditos Hipotecarios...": DoEvents
         Call fs_CargaTipDat(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
      End If
        
      r_str_Fecha3 = Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss")
      pnl_BarTot.FloodPercent = 100
      pnl_BarPro.FloodPercent = 0
      
      cmd_Proces.Enabled = True
      Screen.MousePointer = 0
      MsgBox "Proceso Terminado." & vbCrLf & "Hora 1: " & r_str_Fecha1 & vbCrLf & "Hora 2: " & r_str_Fecha2 & vbCrLf & "Hora 3: " & r_str_Fecha3, vbInformation, modgen_g_str_NomPlt
      pnl_BarTot.FloodPercent = 0
      
   ElseIf CInt(cmb_TipArc.ItemData(cmb_TipArc.ListIndex)) = 1 Then
      
      lbl_NomPro.Caption = "Verificando si existe data para el periodo (Según Estado de Empresa)...": DoEvents
      g_str_Parame = "SELECT NVL(COUNT(*),0) AS TOTREG FROM RCC_CONMEN WHERE CONMEN_PERANO = " & CStr(ipp_PerAno.Text) & " AND CONMEN_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND CONMEN_ESTEMP IS NOT NULL "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      r_lng_Contad = g_rst_Princi!TOTREG
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      If r_lng_Contad > 0 Then
         
         If MsgBox("La información según Estado de Empresa del RCC para este Período ya ha sido cargada. ¿Desea volver a cargar la información?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
            lbl_NomPro.Caption = "Eliminando información según Estado de Empresa...": DoEvents

            'Borra
            pnl_BarPro.FloodPercent = 50
            modprc_g_str_CadEje = "UPDATE RCC_CONMEN SET CONMEN_ESTEMP = NULL WHERE "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "CONMEN_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
            modprc_g_str_CadEje = modprc_g_str_CadEje & "CONMEN_PERANO = " & CStr(ipp_PerAno.Text) & " "
            
            If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 2) Then
               Exit Sub
            End If
            'Carga
            pnl_BarPro.FloodPercent = 100
            lbl_NomPro.Caption = "Proceso carga información según Estado de Empresa...": DoEvents
            Call fs_Carga_EmpSup(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
         End If
      Else
         lbl_NomPro.Caption = "Proceso carga información según Estado de Empresa...": DoEvents
         pnl_BarTot.FloodPercent = 0
         Call fs_Carga_EmpSup(fil_LisArc.Path & "\" & fil_LisArc.FileName, cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text), pnl_BarPro)
      End If
      
      cmd_Proces.Enabled = True
      Screen.MousePointer = 0
      MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub fs_Carga_EmpSup(ByVal p_ArcRCC As String, ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_int_Codigo        As Integer
Dim r_str_TipEmp        As String
Dim r_int_TipEmp        As Integer
Dim r_str_EstEmp        As String
Dim r_str_NomEmp        As String
Dim r_int_Contad        As Integer
Dim r_lng_Contad        As Long

   DoEvents
   p_BarPro.FloodPercent = 0

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=p_ArcRCC
   r_int_FilExc = 2
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
      r_int_Contad = r_int_FilExc - 1
      r_int_FilExc = r_int_FilExc + 1
   Loop
   
   r_int_FilExc = 2
   
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
   
      r_int_Codigo = Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value)                 'Código
      r_str_TipEmp = UCase(Trim(r_obj_Excel.Cells(r_int_FilExc, 2).Value))          'Tipo
      r_int_TipEmp = moddat_gf_TipEmp(r_str_TipEmp)
      r_str_EstEmp = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)                 'Estado
      r_str_NomEmp = UCase(Trim(r_obj_Excel.Cells(r_int_FilExc, 4).Value))          'Nombre
            
      'INGRESA EMPRESA EN CTB_EMPSUP
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "     SELECT NVL(COUNT(*),0) TOTREG "
      g_str_Parame = g_str_Parame & "       FROM CTB_EMPSUP "
      g_str_Parame = g_str_Parame & "      WHERE EMPSUP_CODIGO = " & CInt(r_int_Codigo) & ""
               
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then
      g_rst_Princi.MoveFirst
      r_lng_Contad = g_rst_Princi!TOTREG
      
         If r_lng_Contad = 0 And r_int_Codigo >= 400 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "     INSERT INTO CTB_EMPSUP (EMPSUP_CODIGO, "
            g_str_Parame = g_str_Parame & "                             EMPSUP_NOMBRE, "
            g_str_Parame = g_str_Parame & "                             EMPSUP_NOMCOR, "
            g_str_Parame = g_str_Parame & "                             EMPSUP_TIPENT, "
            g_str_Parame = g_str_Parame & "                             EMPSUP_SITUAC, "
            g_str_Parame = g_str_Parame & "                             SEGUSUCRE, "
            g_str_Parame = g_str_Parame & "                             SEGFECCRE, "
            g_str_Parame = g_str_Parame & "                             SEGHORCRE, "
            g_str_Parame = g_str_Parame & "                             SEGPLTCRE, "
            g_str_Parame = g_str_Parame & "                             SEGTERCRE, "
            g_str_Parame = g_str_Parame & "                             SEGSUCCRE )"
            g_str_Parame = g_str_Parame & "                    VALUES (" & r_int_Codigo & ", "
            g_str_Parame = g_str_Parame & "                           '" & r_str_NomEmp & "', "
            g_str_Parame = g_str_Parame & "                           '" & Mid(r_str_NomEmp, 1, 255) & "', "
            g_str_Parame = g_str_Parame & "                            " & r_int_TipEmp & ", "
            g_str_Parame = g_str_Parame & "                                1, "
            g_str_Parame = g_str_Parame & "                            '" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "                            '" & Format(date, "YYYYMMDD") & "', "
            g_str_Parame = g_str_Parame & "                            '" & Format(Time, "HHMMSS") & "', "
            g_str_Parame = g_str_Parame & "                            '" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "                            '" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "                            '" & modgen_g_str_CodSuc & "') "
                  
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               Exit Sub
            End If
          
            Set g_rst_Genera = Nothing
         End If
      End If
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      'ACTUALIZA ESTADO DE EMPRESA EN RCC_CONMEN
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "     SELECT PARDES_CODITE AS ESTADO_EMPRESA "
      g_str_Parame = g_str_Parame & "       FROM MNT_PARDES "
      g_str_Parame = g_str_Parame & "      WHERE PARDES_CODGRP = 379 "
      g_str_Parame = g_str_Parame & "        AND TRIM (PARDES_DESCRI) = '" & UCase(r_str_EstEmp) & "'"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " UPDATE RCC_CONMEN SET CONMEN_ESTEMP = '" & CInt(g_rst_Princi!ESTADO_EMPRESA) & "'"
         g_str_Parame = g_str_Parame & "  WHERE CONMEN_PERANO = '" & Right("0000" & p_PerAno, 4) & "' AND CONMEN_PERMES = '" & Right("00" & p_PerMes, 2) & "' "
         g_str_Parame = g_str_Parame & "    AND CONMEN_CODEMP = " & r_int_Codigo & ""
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         Set g_rst_Genera = Nothing
      End If
      g_rst_Princi.Close
      
      'ACTUALIZA NOMBRE DE EMPRESA EN RCC_CONMEN
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " UPDATE RCC_CONMEN SET CONMEN_NOMEMP = '" & Mid(r_str_NomEmp, 1, 100) & "'"
      g_str_Parame = g_str_Parame & "  WHERE CONMEN_PERANO = '" & Right("0000" & p_PerAno, 4) & "' AND CONMEN_PERMES = '" & Right("00" & p_PerMes, 2) & "' "
      g_str_Parame = g_str_Parame & "    AND CONMEN_CODEMP = " & r_int_Codigo & " AND CONMEN_NOMEMP IS NULL "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Exit Sub
      End If

      Set g_rst_Genera = Nothing
      
      DoEvents
      p_BarPro.FloodPercent = (r_int_FilExc / r_int_Contad) * 100
      pnl_BarTot.FloodPercent = p_BarPro.FloodPercent
      r_int_FilExc = r_int_FilExc + 1
        
   Loop
   
   Set g_rst_Princi = Nothing
   Set g_rst_Genera = Nothing
   
   r_obj_Excel.Quit
   Set r_obj_Excel = Nothing
End Sub
Private Function moddat_gf_TipEmp(ByVal p_TipoEmp As String) As Integer
   moddat_gf_TipEmp = 0

   g_str_Parame = "SELECT PARDES_CODITE FROM MNT_PARDES WHERE "
   g_str_Parame = g_str_Parame & "PARDES_CODGRP = '263' AND "
   g_str_Parame = g_str_Parame & "PARDES_DESCRI = '" & CStr(p_TipoEmp) & "' AND "
   g_str_Parame = g_str_Parame & "PARDES_SITUAC = 1 "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      moddat_gf_TipEmp = CInt(Trim(g_rst_Princi!PARDES_CODITE & ""))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Private Sub fs_CargaTipDat(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
Dim r_int_TipRep     As Integer
Dim r_int_IndTip     As Integer
Dim r_str_IndTip     As String
Dim r_dbl_Monto      As Double
Dim r_dbl_Total      As Double
Dim r_dbl_MtoSol     As Double
Dim r_dbl_MtoDol     As Double

   DoEvents
   p_BarPro.FloodPercent = 0
   
   'PARA LOS INDICADORES S, V, M90 y J
   For r_int_TipRep = 1 To 2
   
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT A.HIPDET_CODEMP, A.HIPDET_NOMEMP, SUM(A.HIPDET_MTOSOL + A.HIPDET_MTODOL) AS MONTO_SALDOS, "
      g_str_Parame = g_str_Parame & "        SUM(A.HIPDET_MTOSOL) AS MTOSOL_SALDOS, SUM(A.HIPDET_MTODOL) AS MTODOL_SALDOS, "
      g_str_Parame = g_str_Parame & "        E.TOTAL_SALDOS, B.MONTO_VENCIDOS, B.MTOSOL_VENCIDOS, B.MTODOL_VENCIDOS, B.TOTAL_VENCIDOS, C.DIAS_ATRASADOS, "
      g_str_Parame = g_str_Parame & "        D.MONTO_JUDICIAL, D.MTOSOL_JUDICIAL, D.MTODOL_JUDICIAL, D.TOTAL_JUDICIAL "
      
      If r_int_TipRep = 1 Then
         g_str_Parame = g_str_Parame & "       , F.TOTAL_CASTIGADO, F.MONTO_CASTIGADO, F.MTOSOL_CASTIGADO, F.MTODOL_CASTIGADO "
      End If
      
      g_str_Parame = g_str_Parame & "   FROM RCC_HIPDET A "
      g_str_Parame = g_str_Parame & "                LEFT JOIN (SELECT B.HIPDET_CODEMP, B.HIPDET_NOMEMP, COUNT(DISTINCT HIPDET_NOMCLI) AS TOTAL_VENCIDOS, "
      g_str_Parame = g_str_Parame & "                                  SUM(B.HIPDET_MTOSOL + B.HIPDET_MTODOL) AS MONTO_VENCIDOS, "
      g_str_Parame = g_str_Parame & "                                  SUM(B.HIPDET_MTOSOL) AS MTOSOL_VENCIDOS, SUM(B.HIPDET_MTODOL) AS MTODOL_VENCIDOS "
      g_str_Parame = g_str_Parame & "                             FROM RCC_HIPDET B "
      g_str_Parame = g_str_Parame & "                            WHERE B.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' AND B.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "                              AND B.HIPDET_CODEMP > 0"
      g_str_Parame = g_str_Parame & "                              AND B.HIPDET_CODEMP NOT IN (66,191,68,14)"
      g_str_Parame = g_str_Parame & "                              AND TRIM(B.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS'"
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(B.HIPDET_CTACBL,1,4) IN ('1415','1416','1425','1426','1435','1436') "
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(B.HIPDET_CTACBL,1,1) NOT IN ('7','8')"
      
      '14210424','14250424','14260424','14240424'
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & "                           AND (SUBSTR(B.HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                                               '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                                               '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                                               '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                                               '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                                               '14260425','14360425')"
         g_str_Parame = g_str_Parame & "                            OR SUBSTR(B.HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                                              '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                                              '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "                            GROUP BY B.HIPDET_CODEMP, B.HIPDET_NOMEMP) B  ON A.HIPDET_CODEMP = B.HIPDET_CODEMP"
'      g_str_Parame = g_str_Parame & "                                  AND A.HIPDET_NOMEMP = B.HIPDET_NOMEMP "
      g_str_Parame = g_str_Parame & "                LEFT JOIN (SELECT C.HIPDET_CODEMP, C.HIPDET_NOMEMP,"
      g_str_Parame = g_str_Parame & "                                  COUNT(DISTINCT HIPDET_NOMCLI) AS DIAS_ATRASADOS"
      g_str_Parame = g_str_Parame & "                             FROM RCC_HIPDET C"
      g_str_Parame = g_str_Parame & "                            WHERE C.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' AND C.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "                              AND C.HIPDET_CODEMP > 0"
      g_str_Parame = g_str_Parame & "                              AND C.HIPDET_CODEMP NOT IN (66,191,68,14)"
      g_str_Parame = g_str_Parame & "                              AND TRIM(C.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS'"
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(C.HIPDET_CTACBL,1,4) IN ('1415','1416','1425','1426','1435','1436')"
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(C.HIPDET_CTACBL,1,1) NOT IN ('7','8')"
      
      ''14210424','14250424','14260424','14240424',
       If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & "                           AND (SUBSTR(C.HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                                                '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                                                '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                                                '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                                                '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                                                '14260425','14360425')"
         g_str_Parame = g_str_Parame & "                            OR SUBSTR(C.HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                                                '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                                                '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "                              AND C.HIPDET_DIAATR > 90"
      g_str_Parame = g_str_Parame & "                            GROUP BY C.HIPDET_CODEMP, C.HIPDET_NOMEMP) C ON A.HIPDET_CODEMP = C.HIPDET_CODEMP"
'      g_str_Parame = g_str_Parame & "                                  AND A.HIPDET_NOMEMP = C.HIPDET_NOMEMP"
      g_str_Parame = g_str_Parame & "                LEFT JOIN (SELECT D.HIPDET_CODEMP, D.HIPDET_NOMEMP,"
      g_str_Parame = g_str_Parame & "                                  COUNT(DISTINCT HIPDET_NOMCLI) AS TOTAL_JUDICIAL,"
      g_str_Parame = g_str_Parame & "                                  SUM(D.HIPDET_MTOSOL + D.HIPDET_MTODOL) As MONTO_JUDICIAL, "
      g_str_Parame = g_str_Parame & "                                  SUM(D.HIPDET_MTOSOL) AS MTOSOL_JUDICIAL , SUM(D.HIPDET_MTODOL) AS MTODOL_JUDICIAL "
      g_str_Parame = g_str_Parame & "                             FROM RCC_HIPDET D"
      g_str_Parame = g_str_Parame & "                            WHERE D.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' AND D.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "                              AND D.HIPDET_CODEMP > 0"
      g_str_Parame = g_str_Parame & "                              AND D.HIPDET_CODEMP NOT IN (66,191,68,14)"
      g_str_Parame = g_str_Parame & "                              AND TRIM(D.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS'"
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(D.HIPDET_CTACBL,1,4) IN ('1416','1426','1436')"
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(D.HIPDET_CTACBL,1,1) NOT IN ('7','8')"
      
      '14210424','14250424','14260424','14240424',
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & "                           AND (SUBSTR(D.HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                                                '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                                                '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                                                '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                                                '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                                                '14260425','14360425')"
         g_str_Parame = g_str_Parame & "                            OR SUBSTR(D.HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                                                '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                                                '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "                              AND D.HIPDET_DIAATR > 90"
      g_str_Parame = g_str_Parame & "                            GROUP BY D.HIPDET_CODEMP, D.HIPDET_NOMEMP) D  ON A.HIPDET_CODEMP = D.HIPDET_CODEMP"
'      g_str_Parame = g_str_Parame & "                                  AND A.HIPDET_NOMEMP = D.HIPDET_NOMEMP"
      
      g_str_Parame = g_str_Parame & "                LEFT JOIN  (SELECT HIPDET_CODEMP, HIPDET_NOMEMP, SUM(TOTAL_CLIENTES) AS TOTAL_SALDOS "
      g_str_Parame = g_str_Parame & "                              FROM (SELECT E.HIPDET_CODEMP, E.HIPDET_NOMEMP ,(COUNT(DISTINCT HIPDET_NOMCLI)) AS TOTAL_CLIENTES "
      g_str_Parame = g_str_Parame & "                                      FROM RCC_HIPDET E "
      g_str_Parame = g_str_Parame & "                                     WHERE E.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' AND E.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "' "
      g_str_Parame = g_str_Parame & "                                       AND E.HIPDET_CODEMP > 0 "
      g_str_Parame = g_str_Parame & "                                       AND E.HIPDET_CODEMP NOT IN (66,191,68,14) "
      g_str_Parame = g_str_Parame & "                                       AND TRIM(E.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS' "
      g_str_Parame = g_str_Parame & "                                       AND SUBSTR(E.HIPDET_CTACBL,1,1) NOT IN ('7','8') "
            
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & "                           AND (SUBSTR(E.HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                                                '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                                                '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                                                '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                                                '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                                                '14260425','14360425')"
         g_str_Parame = g_str_Parame & "                            OR SUBSTR(E.HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                                                '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                                                '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "                                     GROUP BY E.HIPDET_CODEMP, E.HIPDET_NOMEMP, E.HIPDET_DOCIDE, E.HIPDET_NOMCLI ) X "
      g_str_Parame = g_str_Parame & "                             GROUP BY HIPDET_CODEMP, HIPDET_NOMEMP) E ON A.HIPDET_CODEMP = E.HIPDET_CODEMP "
'      g_str_Parame = g_str_Parame & "                               AND A.HIPDET_NOMEMP = E.HIPDET_NOMEMP "

      If r_int_TipRep = 1 Then
        g_str_Parame = g_str_Parame & "                LEFT JOIN ( SELECT F.HIPDET_CODEMP, F.HIPDET_NOMEMP, COUNT(DISTINCT F.HIPDET_NOMCLI) AS TOTAL_CASTIGADO, "
        g_str_Parame = g_str_Parame & "                                   SUM(F.HIPDET_MTOSOL + F.HIPDET_MTODOL) AS MONTO_CASTIGADO, "
        g_str_Parame = g_str_Parame & "                                   SUM(F.HIPDET_MTOSOL) AS MTOSOL_CASTIGADO, SUM(F.HIPDET_MTODOL) AS MTODOL_CASTIGADO "
        g_str_Parame = g_str_Parame & "                              FROM RCC_HIPDET F "
        g_str_Parame = g_str_Parame & "                             WHERE F.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' "
        g_str_Parame = g_str_Parame & "                               AND F.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "' "
        g_str_Parame = g_str_Parame & "                               AND F.HIPDET_CODEMP > 0 "
        g_str_Parame = g_str_Parame & "                               AND F.HIPDET_CODEMP NOT IN (66,191,68,14) "
        g_str_Parame = g_str_Parame & "                               AND TRIM(F.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS' "
        g_str_Parame = g_str_Parame & "                               AND SUBSTR(F.HIPDET_CTACBL,1,6) IN ('811302','811925','812302','812925') "
        g_str_Parame = g_str_Parame & "                             GROUP BY F.HIPDET_CODEMP, F.HIPDET_NOMEMP) F  ON A.HIPDET_CODEMP = F.HIPDET_CODEMP "
'        g_str_Parame = g_str_Parame & "                                  AND A.HIPDET_NOMEMP = F.HIPDET_NOMEMP "
      End If
      g_str_Parame = g_str_Parame & "  WHERE A.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' "
      g_str_Parame = g_str_Parame & "    AND A.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "    AND A.HIPDET_CODEMP > 0"
      g_str_Parame = g_str_Parame & "    AND A.HIPDET_CODEMP NOT IN (66,191,68,14)"
      g_str_Parame = g_str_Parame & "    AND TRIM(A.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS' "
      g_str_Parame = g_str_Parame & "    AND SUBSTR(A.HIPDET_CTACBL,1,1) NOT IN ('7','8')"
      
      
      '14210424','14250424','14260424','14240424',
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & " AND (SUBSTR(A.HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                      '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                      '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                      '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                      '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                      '14260425','14360425')"
         g_str_Parame = g_str_Parame & "  OR SUBSTR(A.HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                      '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                      '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "  GROUP BY A.HIPDET_CODEMP, A.HIPDET_NOMEMP, B.MONTO_VENCIDOS, B.MTOSOL_VENCIDOS, B.MTODOL_VENCIDOS, B.TOTAL_VENCIDOS, C.DIAS_ATRASADOS, "
      g_str_Parame = g_str_Parame & "           D.MONTO_JUDICIAL, D.MTOSOL_JUDICIAL, D.MTODOL_JUDICIAL, D.TOTAL_JUDICIAL, E.TOTAL_SALDOS "
      
      If r_int_TipRep = 1 Then
         g_str_Parame = g_str_Parame & "           ,F.TOTAL_CASTIGADO, F.MONTO_CASTIGADO, F.MTOSOL_CASTIGADO, F.MTODOL_CASTIGADO "
      End If
      
      g_str_Parame = g_str_Parame & "  ORDER BY MONTO_SALDOS DESC "

      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox g_str_Parame
         Exit Sub
      End If
         
      If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then g_rst_Princi.MoveFirst
      Do Until g_rst_Princi.EOF
         For r_int_IndTip = 1 To 5 '4
            
            If r_int_IndTip = 1 Then r_str_IndTip = "S": r_dbl_Monto = IIf(IsNull(g_rst_Princi!MONTO_SALDOS), 0, g_rst_Princi!MONTO_SALDOS): r_dbl_Total = IIf(IsNull(g_rst_Princi!TOTAL_SALDOS), 0, g_rst_Princi!TOTAL_SALDOS): r_dbl_MtoSol = IIf(IsNull(g_rst_Princi!MTOSOL_SALDOS), 0, g_rst_Princi!MTOSOL_SALDOS): r_dbl_MtoDol = IIf(IsNull(g_rst_Princi!MTODOL_SALDOS), 0, g_rst_Princi!MTODOL_SALDOS)
            If r_int_IndTip = 2 Then r_str_IndTip = "V": r_dbl_Monto = IIf(IsNull(g_rst_Princi!MONTO_VENCIDOS), 0, g_rst_Princi!MONTO_VENCIDOS): r_dbl_Total = IIf(IsNull(g_rst_Princi!TOTAL_VENCIDOS), 0, g_rst_Princi!TOTAL_VENCIDOS): r_dbl_MtoSol = IIf(IsNull(g_rst_Princi!MTOSOL_VENCIDOS), 0, g_rst_Princi!MTOSOL_VENCIDOS): r_dbl_MtoDol = IIf(IsNull(g_rst_Princi!MTODOL_VENCIDOS), 0, g_rst_Princi!MTODOL_VENCIDOS)
            If r_int_IndTip = 3 Then r_str_IndTip = "J": r_dbl_Monto = IIf(IsNull(g_rst_Princi!MONTO_JUDICIAL), 0, g_rst_Princi!MONTO_JUDICIAL): r_dbl_Total = IIf(IsNull(g_rst_Princi!TOTAL_JUDICIAL), 0, g_rst_Princi!TOTAL_JUDICIAL): r_dbl_MtoSol = IIf(IsNull(g_rst_Princi!MTOSOL_JUDICIAL), 0, g_rst_Princi!MTOSOL_JUDICIAL): r_dbl_MtoDol = IIf(IsNull(g_rst_Princi!MTODOL_JUDICIAL), 0, g_rst_Princi!MTODOL_JUDICIAL)
            
            If r_int_TipRep = 1 Then
               If r_int_IndTip = 4 Then r_str_IndTip = "C": r_dbl_Monto = IIf(IsNull(g_rst_Princi!MONTO_CASTIGADO), 0, g_rst_Princi!MONTO_CASTIGADO): r_dbl_Total = IIf(IsNull(g_rst_Princi!TOTAL_CASTIGADO), 0, g_rst_Princi!TOTAL_CASTIGADO): r_dbl_MtoSol = IIf(IsNull(g_rst_Princi!MTOSOL_CASTIGADO), 0, g_rst_Princi!MTOSOL_CASTIGADO): r_dbl_MtoDol = IIf(IsNull(g_rst_Princi!MTODOL_CASTIGADO), 0, g_rst_Princi!MTODOL_CASTIGADO)
            End If
            
            If r_int_IndTip = 5 Then r_str_IndTip = "M90":  r_dbl_Total = IIf(IsNull(g_rst_Princi!DIAS_ATRASADOS), 0, g_rst_Princi!DIAS_ATRASADOS)
            
            If r_dbl_Monto > 0 Or r_dbl_Total > 0 Then
               g_str_Parame = ""
               g_str_Parame = g_str_Parame & "INSERT INTO RCC_CONMEN("
               g_str_Parame = g_str_Parame & "CONMEN_PERANO, "
               g_str_Parame = g_str_Parame & "CONMEN_PERMES, "
               g_str_Parame = g_str_Parame & "CONMEN_CODEMP, "
               g_str_Parame = g_str_Parame & "CONMEN_TIPREP, "
               g_str_Parame = g_str_Parame & "CONMEN_INDTIP, "
               g_str_Parame = g_str_Parame & "CONMEN_NOMEMP, "
               g_str_Parame = g_str_Parame & "CONMEN_MONTOT, "
               g_str_Parame = g_str_Parame & "CONMEN_NUMTOT, "
               g_str_Parame = g_str_Parame & "CONMEN_MTOSOL, "
               g_str_Parame = g_str_Parame & "CONMEN_MTODOL) "
               g_str_Parame = g_str_Parame & "VALUES ("
               g_str_Parame = g_str_Parame & "" & Format(p_PerAno, "0000") & ", "
               g_str_Parame = g_str_Parame & "" & Format(p_PerMes, "00") & ", "
               g_str_Parame = g_str_Parame & "" & Trim(g_rst_Princi!HIPDET_CODEMP) & " , "
               g_str_Parame = g_str_Parame & "" & Trim(r_int_TipRep) & " , "
               g_str_Parame = g_str_Parame & "'" & Trim(r_str_IndTip) & "' , "
               g_str_Parame = g_str_Parame & "'" & gf_Buscar_NomEmp(Trim(g_rst_Princi!HIPDET_CODEMP)) & "' , " 'Trim(g_rst_Princi!HIPDET_NOMEMP)
               g_str_Parame = g_str_Parame & IIf(r_dbl_Monto = 0, "NULL", r_dbl_Monto) & ", "   'Format(Trim(r_dbl_Monto), "########0.00")
               g_str_Parame = g_str_Parame & r_dbl_Total & " , "                                'Format(Trim(r_dbl_Total), "########0.00")
               g_str_Parame = g_str_Parame & IIf(r_dbl_Monto = 0, "NULL", r_dbl_MtoSol) & " , "
               g_str_Parame = g_str_Parame & IIf(r_dbl_Monto = 0, "NULL", r_dbl_MtoDol) & ")"
               
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                  MsgBox g_str_Parame
                  Exit Sub
               End If
               
               r_dbl_Monto = 0
               r_dbl_Total = 0
            End If
         Next r_int_IndTip
         g_rst_Princi.MoveNext
      Loop
      DoEvents
      p_BarPro.FloodPercent = p_BarPro.FloodPercent + 15
   Next r_int_TipRep
   
   g_rst_Princi.Close

   'PARA EL INDICADOR HT
   For r_int_TipRep = 1 To 2
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "   SELECT HIPDET_CODEMP ,HIPDET_NOMEMP, SUM(HIPDET_MTOSOL+HIPDET_MTODOL) AS MTO_ACT_TOTAL, SUM(HIPDET_MTOSOL) AS MONTO_SOLES, SUM(HIPDET_MTODOL) AS MONTO_DOLARES "
      g_str_Parame = g_str_Parame & "     FROM RCC_HIPDET"
      g_str_Parame = g_str_Parame & "    WHERE HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' "
      g_str_Parame = g_str_Parame & "      AND HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "      AND HIPDET_CODEMP > 0 "
      g_str_Parame = g_str_Parame & "      AND HIPDET_TIPDEU = 'CREDITOS HIPOTECARIOS' "
      
      '14210424','14250424','14260424','14240424',
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & "   AND SUBSTR(HIPDET_CTACBL,1,1) NOT IN ('7','8')"
         g_str_Parame = g_str_Parame & "   AND (SUBSTR(HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                      '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                      '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                      '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                      '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                      '14260425','14360425')"
         g_str_Parame = g_str_Parame & "    OR SUBSTR(HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                      '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                      '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "      AND HIPDET_CTACBL > 0 "
      g_str_Parame = g_str_Parame & "      AND HIPDET_NUMITE > 0 "
      g_str_Parame = g_str_Parame & "    GROUP BY HIPDET_CODEMP,HIPDET_NOMEMP,HIPDET_PERANO,HIPDET_PERMES "
      
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox g_str_Parame
         Exit Sub
      End If
      
      r_dbl_MtoSol = 0
      r_dbl_MtoDol = 0
      If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then g_rst_Princi.MoveFirst
      Do Until g_rst_Princi.EOF
         If Not IsNull(g_rst_Princi!MTO_ACT_TOTAL) Then
            r_dbl_Monto = g_rst_Princi!MTO_ACT_TOTAL
            r_dbl_MtoSol = g_rst_Princi!MONTO_SOLES
            r_dbl_MtoDol = g_rst_Princi!MONTO_DOLARES
            r_str_IndTip = "HT"
            r_dbl_Total = 0
         End If
         
         If r_dbl_Monto > 0 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO RCC_CONMEN("
            g_str_Parame = g_str_Parame & "CONMEN_PERANO, "
            g_str_Parame = g_str_Parame & "CONMEN_PERMES, "
            g_str_Parame = g_str_Parame & "CONMEN_CODEMP, "
            g_str_Parame = g_str_Parame & "CONMEN_TIPREP, "
            g_str_Parame = g_str_Parame & "CONMEN_INDTIP, "
            g_str_Parame = g_str_Parame & "CONMEN_NOMEMP, "
            g_str_Parame = g_str_Parame & "CONMEN_MONTOT, "
            g_str_Parame = g_str_Parame & "CONMEN_NUMTOT, "
            g_str_Parame = g_str_Parame & "CONMEN_MTOSOL, "
            g_str_Parame = g_str_Parame & "CONMEN_MTODOL) "
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & "" & Format(p_PerAno, "0000") & ", "
            g_str_Parame = g_str_Parame & "" & Format(p_PerMes, "00") & ", "
            g_str_Parame = g_str_Parame & "" & Trim(g_rst_Princi!HIPDET_CODEMP) & " , "
            g_str_Parame = g_str_Parame & "" & Trim(r_int_TipRep) & " , "
            g_str_Parame = g_str_Parame & "'" & Trim(r_str_IndTip) & "' , "
            g_str_Parame = g_str_Parame & "'" & gf_Buscar_NomEmp(Trim(g_rst_Princi!HIPDET_CODEMP)) & "' , " 'Trim(g_rst_Princi!HIPDET_NOMEMP)
            g_str_Parame = g_str_Parame & r_dbl_Monto & ", "                                 'Format(Trim(r_dbl_Monto), "########0.00")
            g_str_Parame = g_str_Parame & IIf(r_dbl_Total = 0, "NULL", r_dbl_Total) & " , "  'Format(Trim(r_dbl_Total), "########0.00")
            g_str_Parame = g_str_Parame & r_dbl_MtoSol & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoDol & ")"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               MsgBox g_str_Parame
               Exit Sub
            End If
            
            r_dbl_Monto = 0
            r_dbl_Total = 0
            r_dbl_MtoSol = 0
            r_dbl_MtoDol = 0
         End If
         g_rst_Princi.MoveNext
      Loop
      DoEvents
      p_BarPro.FloodPercent = p_BarPro.FloodPercent + 15
   Next r_int_TipRep
   
   g_rst_Princi.Close
   pnl_BarTot.FloodPercent = 90
   
   'PARA EL INDICADOR HP
   r_dbl_MtoSol = 0
   r_dbl_MtoDol = 0
   
   For r_int_TipRep = 1 To 2
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT HIPDET_CODEMP,HIPDET_NOMEMP, SUM(HIPDET_MTOSOL+HIPDET_MTODOL) AS MTO_ACT_PESADO, SUM(HIPDET_MTOSOL) AS MONTO_SOLES, SUM(HIPDET_MTODOL) AS MONTO_DOLARES "
      g_str_Parame = g_str_Parame & "   FROM RCC_HIPDET"
      g_str_Parame = g_str_Parame & "  WHERE HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "'"
      g_str_Parame = g_str_Parame & "    AND HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "    AND HIPDET_CODEMP > 0"
      g_str_Parame = g_str_Parame & "    AND HIPDET_TIPDEU = 'CREDITOS HIPOTECARIOS'"
      
      ''14210424','14250424','14260424','14240424',
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & " AND SUBSTR(HIPDET_CTACBL,1,1) NOT IN ('7','8')"
         g_str_Parame = g_str_Parame & " AND (SUBSTR(HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                    '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                    '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                    '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                    '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                    '14260425','14360425')"
         g_str_Parame = g_str_Parame & "     OR SUBSTR(HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                       '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                       '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "    AND HIPDET_CTACBL > 0"
      g_str_Parame = g_str_Parame & "    AND HIPDET_NUMITE > 0"
      g_str_Parame = g_str_Parame & "    AND (HIPDET_CLASIF = 'DEFICIENTE' OR HIPDET_CLASIF = 'DUDOSO' OR HIPDET_CLASIF = 'PERDIDA')"
      g_str_Parame = g_str_Parame & "  GROUP BY HIPDET_CODEMP,HIPDET_NOMEMP "
       
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox g_str_Parame
         Exit Sub
      End If
      
      If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then g_rst_Princi.MoveFirst
      
      Do Until g_rst_Princi.EOF
      
         If Not IsNull(g_rst_Princi!MTO_ACT_PESADO) Then
            r_dbl_Monto = g_rst_Princi!MTO_ACT_PESADO
            r_dbl_MtoSol = g_rst_Princi!MONTO_SOLES
            r_dbl_MtoDol = g_rst_Princi!MONTO_DOLARES
            r_str_IndTip = "HP"
            r_dbl_Total = 0
         End If
         If r_dbl_Monto > 0 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO RCC_CONMEN("
            g_str_Parame = g_str_Parame & "CONMEN_PERANO, "
            g_str_Parame = g_str_Parame & "CONMEN_PERMES, "
            g_str_Parame = g_str_Parame & "CONMEN_CODEMP, "
            g_str_Parame = g_str_Parame & "CONMEN_TIPREP, "
            g_str_Parame = g_str_Parame & "CONMEN_INDTIP, "
            g_str_Parame = g_str_Parame & "CONMEN_NOMEMP, "
            g_str_Parame = g_str_Parame & "CONMEN_MONTOT, "
            g_str_Parame = g_str_Parame & "CONMEN_NUMTOT, "
            g_str_Parame = g_str_Parame & "CONMEN_MTOSOL, "
            g_str_Parame = g_str_Parame & "CONMEN_MTODOL) "
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & "" & Format(p_PerAno, "0000") & ", "
            g_str_Parame = g_str_Parame & "" & Format(p_PerMes, "00") & ", "
            g_str_Parame = g_str_Parame & "" & Trim(g_rst_Princi!HIPDET_CODEMP) & " , "
            g_str_Parame = g_str_Parame & "" & Trim(r_int_TipRep) & " , "
            g_str_Parame = g_str_Parame & "'" & Trim(r_str_IndTip) & "' , "
            g_str_Parame = g_str_Parame & "'" & gf_Buscar_NomEmp(Trim(g_rst_Princi!HIPDET_CODEMP)) & "' , " 'Trim(g_rst_Princi!HIPDET_NOMEMP)
            g_str_Parame = g_str_Parame & r_dbl_Monto & ", "                                 'Format(Trim(r_dbl_Monto), "########0.00")
            g_str_Parame = g_str_Parame & IIf(r_dbl_Total = 0, "NULL", r_dbl_Total) & " , "  'Format(Trim(r_dbl_Total), "########0.00")
            g_str_Parame = g_str_Parame & r_dbl_MtoSol & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoDol & ")"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               MsgBox g_str_Parame
               Exit Sub
            End If
            
            r_dbl_Monto = 0
            r_dbl_Total = 0
            r_dbl_MtoSol = 0
            r_dbl_MtoDol = 0
         End If
         
         g_rst_Princi.MoveNext
      Loop
      DoEvents
      p_BarPro.FloodPercent = p_BarPro.FloodPercent + 15
   Next r_int_TipRep
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   DoEvents
   p_BarPro.FloodPercent = 100
End Sub

Private Sub fs_CargaTipDatOld(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, Optional p_BarPro As SSPanel)
Dim r_int_TipRep     As Integer
Dim r_int_IndTip     As Integer
Dim r_str_IndTip     As String
Dim r_dbl_Monto      As Double
Dim r_dbl_Total      As Double
Dim r_dbl_MtoSol     As Double
Dim r_dbl_MtoDol     As Double

   DoEvents
   p_BarPro.FloodPercent = 0
   
   'PARA LOS INDICADORES S, V, M90 y J
   For r_int_TipRep = 1 To 2
   
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT A.HIPDET_CODEMP, A.HIPDET_NOMEMP, SUM(A.HIPDET_MTOSOL + A.HIPDET_MTODOL) AS MONTO_SALDOS, "
      g_str_Parame = g_str_Parame & "        SUM(A.HIPDET_MTOSOL) AS MTOSOL_SALDOS, SUM(A.HIPDET_MTODOL) AS MTODOL_SALDOS, "
      g_str_Parame = g_str_Parame & "        E.TOTAL_SALDOS, B.MONTO_VENCIDOS, B.MTOSOL_VENCIDOS, B.MTODOL_VENCIDOS, B.TOTAL_VENCIDOS, C.DIAS_ATRASADOS, "
      g_str_Parame = g_str_Parame & "        D.MONTO_JUDICIAL, D.MTOSOL_JUDICIAL, D.MTODOL_JUDICIAL, D.TOTAL_JUDICIAL "
      g_str_Parame = g_str_Parame & "   FROM RCC_HIPDET A "
      g_str_Parame = g_str_Parame & "                LEFT JOIN (SELECT B.HIPDET_CODEMP, B.HIPDET_NOMEMP, COUNT(DISTINCT HIPDET_NOMCLI) AS TOTAL_VENCIDOS, "
      g_str_Parame = g_str_Parame & "                                  SUM(B.HIPDET_MTOSOL + B.HIPDET_MTODOL) AS MONTO_VENCIDOS, "
      g_str_Parame = g_str_Parame & "                                  SUM(B.HIPDET_MTOSOL) AS MTOSOL_VENCIDOS, SUM(B.HIPDET_MTODOL) AS MTODOL_VENCIDOS "
      g_str_Parame = g_str_Parame & "                             FROM RCC_HIPDET B "
      g_str_Parame = g_str_Parame & "                            WHERE B.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' AND B.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "                              AND B.HIPDET_CODEMP < 300"
      g_str_Parame = g_str_Parame & "                              AND B.HIPDET_CODEMP NOT IN (66,191,68,14)"
      g_str_Parame = g_str_Parame & "                              AND TRIM(B.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS'"
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(B.HIPDET_CTACBL,1,4) IN ('1415','1416','1425','1426','1435','1436') "
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(B.HIPDET_CTACBL,1,1) NOT IN ('7','8')"
      
      '14210424','14250424','14260424','14240424'
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & "                           AND (SUBSTR(B.HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                                               '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                                               '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                                               '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                                               '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                                               '14260425','14360425')"
         g_str_Parame = g_str_Parame & "                            OR SUBSTR(B.HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                                              '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                                              '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "                            GROUP BY B.HIPDET_CODEMP, B.HIPDET_NOMEMP) B  ON A.HIPDET_CODEMP = B.HIPDET_CODEMP"
      g_str_Parame = g_str_Parame & "                                  AND A.HIPDET_NOMEMP = B.HIPDET_NOMEMP "
      g_str_Parame = g_str_Parame & "                LEFT JOIN (SELECT C.HIPDET_CODEMP, C.HIPDET_NOMEMP,"
      g_str_Parame = g_str_Parame & "                                  COUNT(DISTINCT HIPDET_NOMCLI) AS DIAS_ATRASADOS"
      g_str_Parame = g_str_Parame & "                             FROM RCC_HIPDET C"
      g_str_Parame = g_str_Parame & "                            WHERE C.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' AND C.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "                              AND C.HIPDET_CODEMP < 300"
      g_str_Parame = g_str_Parame & "                              AND C.HIPDET_CODEMP NOT IN (66,191,68,14)"
      g_str_Parame = g_str_Parame & "                              AND TRIM(C.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS'"
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(C.HIPDET_CTACBL,1,4) IN ('1415','1416','1425','1426','1435','1436')"
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(C.HIPDET_CTACBL,1,1) NOT IN ('7','8')"
      
      ''14210424','14250424','14260424','14240424',
       If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & "                           AND (SUBSTR(C.HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                                                '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                                                '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                                                '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                                                '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                                                '14260425','14360425')"
         g_str_Parame = g_str_Parame & "                            OR SUBSTR(C.HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                                                '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                                                '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "                              AND C.HIPDET_DIAATR > 90"
      g_str_Parame = g_str_Parame & "                            GROUP BY C.HIPDET_CODEMP, C.HIPDET_NOMEMP) C ON A.HIPDET_CODEMP = C.HIPDET_CODEMP"
      g_str_Parame = g_str_Parame & "                                  AND A.HIPDET_NOMEMP = C.HIPDET_NOMEMP"
      g_str_Parame = g_str_Parame & "                LEFT JOIN (SELECT D.HIPDET_CODEMP, D.HIPDET_NOMEMP,"
      g_str_Parame = g_str_Parame & "                                  COUNT(DISTINCT HIPDET_NOMCLI) AS TOTAL_JUDICIAL,"
      g_str_Parame = g_str_Parame & "                                  SUM(D.HIPDET_MTOSOL + D.HIPDET_MTODOL) As MONTO_JUDICIAL, "
      g_str_Parame = g_str_Parame & "                                  SUM(D.HIPDET_MTOSOL) AS MTOSOL_JUDICIAL , SUM(D.HIPDET_MTODOL) AS MTODOL_JUDICIAL "
      g_str_Parame = g_str_Parame & "                             FROM RCC_HIPDET D"
      g_str_Parame = g_str_Parame & "                            WHERE D.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' AND D.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "                              AND D.HIPDET_CODEMP < 300"
      g_str_Parame = g_str_Parame & "                              AND D.HIPDET_CODEMP NOT IN (66,191,68,14)"
      g_str_Parame = g_str_Parame & "                              AND TRIM(D.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS'"
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(D.HIPDET_CTACBL,1,4) IN ('1416','1426','1436')"
      g_str_Parame = g_str_Parame & "                              AND SUBSTR(D.HIPDET_CTACBL,1,1) NOT IN ('7','8')"
      
      '14210424','14250424','14260424','14240424',
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & "                           AND (SUBSTR(D.HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                                                '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                                                '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                                                '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                                                '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                                                '14260425','14360425')"
         g_str_Parame = g_str_Parame & "                            OR SUBSTR(D.HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                                                '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                                                '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "                              AND D.HIPDET_DIAATR > 90"
      g_str_Parame = g_str_Parame & "                            GROUP BY D.HIPDET_CODEMP, D.HIPDET_NOMEMP) D  ON A.HIPDET_CODEMP = D.HIPDET_CODEMP"
      g_str_Parame = g_str_Parame & "                                  AND A.HIPDET_NOMEMP = D.HIPDET_NOMEMP"
      
      
      g_str_Parame = g_str_Parame & "                LEFT JOIN  (SELECT HIPDET_CODEMP, HIPDET_NOMEMP, SUM(TOTAL_CLIENTES) AS TOTAL_SALDOS "
      g_str_Parame = g_str_Parame & "                              FROM (SELECT E.HIPDET_CODEMP, E.HIPDET_NOMEMP ,(COUNT(DISTINCT HIPDET_NOMCLI)) AS TOTAL_CLIENTES "
      g_str_Parame = g_str_Parame & "                                      FROM RCC_HIPDET E "
      g_str_Parame = g_str_Parame & "                                     WHERE E.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' AND E.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "' "
      g_str_Parame = g_str_Parame & "                                       AND E.HIPDET_CODEMP < 300 "
      g_str_Parame = g_str_Parame & "                                       AND E.HIPDET_CODEMP NOT IN (66,191,68,14) "
      g_str_Parame = g_str_Parame & "                                       AND TRIM(E.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS' "
      g_str_Parame = g_str_Parame & "                                       AND SUBSTR(E.HIPDET_CTACBL,1,1) NOT IN ('7','8') "
            
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & "                           AND (SUBSTR(E.HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                                                '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                                                '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                                                '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                                                '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                                                '14260425','14360425')"
         g_str_Parame = g_str_Parame & "                            OR SUBSTR(E.HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                                                '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                                                '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "                                     GROUP BY E.HIPDET_CODEMP, E.HIPDET_NOMEMP, E.HIPDET_DOCIDE, E.HIPDET_NOMCLI ) X "
      g_str_Parame = g_str_Parame & "                             GROUP BY HIPDET_CODEMP, HIPDET_NOMEMP) E ON A.HIPDET_CODEMP = E.HIPDET_CODEMP "
      g_str_Parame = g_str_Parame & "                               AND A.HIPDET_NOMEMP = E.HIPDET_NOMEMP "


      g_str_Parame = g_str_Parame & "  WHERE A.HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' "
      g_str_Parame = g_str_Parame & "    AND A.HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "    AND A.HIPDET_CODEMP < 300"
      g_str_Parame = g_str_Parame & "    AND A.HIPDET_CODEMP NOT IN (66,191,68,14)"
      g_str_Parame = g_str_Parame & "    AND TRIM(A.HIPDET_TIPDEU) = 'CREDITOS HIPOTECARIOS' "
      g_str_Parame = g_str_Parame & "    AND SUBSTR(A.HIPDET_CTACBL,1,1) NOT IN ('7','8')"
      
      
      '14210424','14250424','14260424','14240424',
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & " AND (SUBSTR(A.HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                      '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                      '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                      '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                      '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                      '14260425','14360425')"
         g_str_Parame = g_str_Parame & "  OR SUBSTR(A.HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                      '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                      '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "  GROUP BY A.HIPDET_CODEMP, A.HIPDET_NOMEMP, B.MONTO_VENCIDOS, B.MTOSOL_VENCIDOS, B.MTODOL_VENCIDOS, B.TOTAL_VENCIDOS, C.DIAS_ATRASADOS, "
      g_str_Parame = g_str_Parame & "           D.MONTO_JUDICIAL, D.MTOSOL_JUDICIAL, D.MTODOL_JUDICIAL, D.TOTAL_JUDICIAL, E.TOTAL_SALDOS "
      g_str_Parame = g_str_Parame & "  ORDER BY MONTO_SALDOS DESC "

      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
      End If
         
      If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then g_rst_Princi.MoveFirst
      Do Until g_rst_Princi.EOF
         For r_int_IndTip = 1 To 4
            
            If r_int_IndTip = 1 Then r_str_IndTip = "S": r_dbl_Monto = IIf(IsNull(g_rst_Princi!MONTO_SALDOS), 0, g_rst_Princi!MONTO_SALDOS): r_dbl_Total = IIf(IsNull(g_rst_Princi!TOTAL_SALDOS), 0, g_rst_Princi!TOTAL_SALDOS): r_dbl_MtoSol = IIf(IsNull(g_rst_Princi!MTOSOL_SALDOS), 0, g_rst_Princi!MTOSOL_SALDOS): r_dbl_MtoDol = IIf(IsNull(g_rst_Princi!MTODOL_SALDOS), 0, g_rst_Princi!MTODOL_SALDOS)
            If r_int_IndTip = 2 Then r_str_IndTip = "V": r_dbl_Monto = IIf(IsNull(g_rst_Princi!MONTO_VENCIDOS), 0, g_rst_Princi!MONTO_VENCIDOS): r_dbl_Total = IIf(IsNull(g_rst_Princi!TOTAL_VENCIDOS), 0, g_rst_Princi!TOTAL_VENCIDOS): r_dbl_MtoSol = IIf(IsNull(g_rst_Princi!MTOSOL_VENCIDOS), 0, g_rst_Princi!MTOSOL_VENCIDOS): r_dbl_MtoDol = IIf(IsNull(g_rst_Princi!MTODOL_VENCIDOS), 0, g_rst_Princi!MTODOL_VENCIDOS)
            If r_int_IndTip = 3 Then r_str_IndTip = "J": r_dbl_Monto = IIf(IsNull(g_rst_Princi!MONTO_JUDICIAL), 0, g_rst_Princi!MONTO_JUDICIAL): r_dbl_Total = IIf(IsNull(g_rst_Princi!TOTAL_JUDICIAL), 0, g_rst_Princi!TOTAL_JUDICIAL): r_dbl_MtoSol = IIf(IsNull(g_rst_Princi!MTOSOL_JUDICIAL), 0, g_rst_Princi!MTOSOL_JUDICIAL): r_dbl_MtoDol = IIf(IsNull(g_rst_Princi!MTODOL_JUDICIAL), 0, g_rst_Princi!MTODOL_JUDICIAL)
            If r_int_IndTip = 4 Then r_str_IndTip = "M90":  r_dbl_Total = IIf(IsNull(g_rst_Princi!DIAS_ATRASADOS), 0, g_rst_Princi!DIAS_ATRASADOS)
            
            If r_dbl_Monto > 0 Or r_dbl_Total > 0 Then
               g_str_Parame = ""
               g_str_Parame = g_str_Parame & "INSERT INTO RCC_CONMEN("
               g_str_Parame = g_str_Parame & "CONMEN_PERANO, "
               g_str_Parame = g_str_Parame & "CONMEN_PERMES, "
               g_str_Parame = g_str_Parame & "CONMEN_CODEMP, "
               g_str_Parame = g_str_Parame & "CONMEN_TIPREP, "
               g_str_Parame = g_str_Parame & "CONMEN_INDTIP, "
               g_str_Parame = g_str_Parame & "CONMEN_NOMEMP, "
               g_str_Parame = g_str_Parame & "CONMEN_MONTOT, "
               g_str_Parame = g_str_Parame & "CONMEN_NUMTOT, "
               g_str_Parame = g_str_Parame & "CONMEN_MTOSOL, "
               g_str_Parame = g_str_Parame & "CONMEN_MTODOL) "
               g_str_Parame = g_str_Parame & "VALUES ("
               g_str_Parame = g_str_Parame & "'" & Format(p_PerAno, "0000") & "', "
               g_str_Parame = g_str_Parame & "'" & Format(p_PerMes, "00") & "', "
               g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!HIPDET_CODEMP) & "' , "
               g_str_Parame = g_str_Parame & "'" & Trim(r_int_TipRep) & "' , "
               g_str_Parame = g_str_Parame & "'" & Trim(r_str_IndTip) & "' , "
               g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!HIPDET_NOMEMP) & "' , "
               g_str_Parame = g_str_Parame & IIf(r_dbl_Monto = 0, "NULL", r_dbl_Monto) & ", "   'Format(Trim(r_dbl_Monto), "########0.00")
               g_str_Parame = g_str_Parame & r_dbl_Total & " , "                                'Format(Trim(r_dbl_Total), "########0.00")
               g_str_Parame = g_str_Parame & IIf(r_dbl_Monto = 0, "NULL", r_dbl_MtoSol) & " , "
               g_str_Parame = g_str_Parame & IIf(r_dbl_Monto = 0, "NULL", r_dbl_MtoDol) & ")"
               
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
                  Exit Sub
               End If
               r_dbl_Monto = 0
               r_dbl_Total = 0
            End If
         Next r_int_IndTip
         g_rst_Princi.MoveNext
      Loop
      DoEvents
      p_BarPro.FloodPercent = p_BarPro.FloodPercent + 15
   Next r_int_TipRep
   
   g_rst_Princi.Close
   
   'PARA EL INDICADOR HT
   For r_int_TipRep = 1 To 2
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "   SELECT HIPDET_CODEMP ,HIPDET_NOMEMP, SUM(HIPDET_MTOSOL+HIPDET_MTODOL) AS MTO_ACT_TOTAL, SUM(HIPDET_MTOSOL) AS MONTO_SOLES, SUM(HIPDET_MTODOL) AS MONTO_DOLARES "
      g_str_Parame = g_str_Parame & "     FROM RCC_HIPDET"
      g_str_Parame = g_str_Parame & "    WHERE HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "' "
      g_str_Parame = g_str_Parame & "      AND HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "      AND HIPDET_CODEMP > 0 "
      g_str_Parame = g_str_Parame & "      AND HIPDET_TIPDEU = 'CREDITOS HIPOTECARIOS' "
      
      '14210424','14250424','14260424','14240424',
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & "   AND SUBSTR(HIPDET_CTACBL,1,1) NOT IN ('7','8')"
         g_str_Parame = g_str_Parame & "   AND (SUBSTR(HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                      '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                      '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                      '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                      '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                      '14260425','14360425')"
         g_str_Parame = g_str_Parame & "    OR SUBSTR(HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                      '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                      '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "      AND HIPDET_CTACBL > 0 "
      g_str_Parame = g_str_Parame & "      AND HIPDET_NUMITE > 0 "
      g_str_Parame = g_str_Parame & "    GROUP BY HIPDET_CODEMP,HIPDET_NOMEMP,HIPDET_PERANO,HIPDET_PERMES "
      
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
      End If
      r_dbl_MtoSol = 0
      r_dbl_MtoDol = 0
      If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then g_rst_Princi.MoveFirst
      Do Until g_rst_Princi.EOF
         If Not IsNull(g_rst_Princi!MTO_ACT_TOTAL) Then
            r_dbl_Monto = g_rst_Princi!MTO_ACT_TOTAL
            r_dbl_MtoSol = g_rst_Princi!MONTO_SOLES
            r_dbl_MtoDol = g_rst_Princi!MONTO_DOLARES
            r_str_IndTip = "HT"
            r_dbl_Total = 0
         End If
         
         If r_dbl_Monto > 0 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO RCC_CONMEN("
            g_str_Parame = g_str_Parame & "CONMEN_PERANO, "
            g_str_Parame = g_str_Parame & "CONMEN_PERMES, "
            g_str_Parame = g_str_Parame & "CONMEN_CODEMP, "
            g_str_Parame = g_str_Parame & "CONMEN_TIPREP, "
            g_str_Parame = g_str_Parame & "CONMEN_INDTIP, "
            g_str_Parame = g_str_Parame & "CONMEN_NOMEMP, "
            g_str_Parame = g_str_Parame & "CONMEN_MONTOT, "
            g_str_Parame = g_str_Parame & "CONMEN_NUMTOT, "
            g_str_Parame = g_str_Parame & "CONMEN_MTOSOL, "
            g_str_Parame = g_str_Parame & "CONMEN_MTODOL) "
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & "'" & Format(p_PerAno, "0000") & "', "
            g_str_Parame = g_str_Parame & "'" & Format(p_PerMes, "00") & "', "
            g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!HIPDET_CODEMP) & "' , "
            g_str_Parame = g_str_Parame & "'" & Trim(r_int_TipRep) & "' , "
            g_str_Parame = g_str_Parame & "'" & Trim(r_str_IndTip) & "' , "
            g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!HIPDET_NOMEMP) & "' , "
            g_str_Parame = g_str_Parame & r_dbl_Monto & ", "                                 'Format(Trim(r_dbl_Monto), "########0.00")
            g_str_Parame = g_str_Parame & IIf(r_dbl_Total = 0, "NULL", r_dbl_Total) & " , "  'Format(Trim(r_dbl_Total), "########0.00")
            g_str_Parame = g_str_Parame & r_dbl_MtoSol & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoDol & ")"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               Exit Sub
            End If
            r_dbl_Monto = 0
            r_dbl_Total = 0
            r_dbl_MtoSol = 0
            r_dbl_MtoDol = 0
         End If
         g_rst_Princi.MoveNext
      Loop
      DoEvents
      p_BarPro.FloodPercent = p_BarPro.FloodPercent + 15
   Next r_int_TipRep
   
   g_rst_Princi.Close
   pnl_BarTot.FloodPercent = 90
   
   'PARA EL INDICADOR HP
   r_dbl_MtoSol = 0
   r_dbl_MtoDol = 0
   
   For r_int_TipRep = 1 To 2
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT HIPDET_CODEMP,HIPDET_NOMEMP, SUM(HIPDET_MTOSOL+HIPDET_MTODOL) AS MTO_ACT_PESADO, SUM(HIPDET_MTOSOL) AS MONTO_SOLES, SUM(HIPDET_MTODOL) AS MONTO_DOLARES "
      g_str_Parame = g_str_Parame & "   FROM RCC_HIPDET"
      g_str_Parame = g_str_Parame & "  WHERE HIPDET_PERANO = '" & Right("0000" & p_PerAno, 4) & "'"
      g_str_Parame = g_str_Parame & "    AND HIPDET_PERMES = '" & Right("00" & p_PerMes, 2) & "'"
      g_str_Parame = g_str_Parame & "    AND HIPDET_CODEMP > 0"
      g_str_Parame = g_str_Parame & "    AND HIPDET_TIPDEU = 'CREDITOS HIPOTECARIOS'"
      
      ''14210424','14250424','14260424','14240424',
      If r_int_TipRep = 2 Then
         g_str_Parame = g_str_Parame & " AND SUBSTR(HIPDET_CTACBL,1,1) NOT IN ('7','8')"
         g_str_Parame = g_str_Parame & " AND (SUBSTR(HIPDET_CTACBL,1,8) IN ('14110423','14210423','14310423','14110424','14310424','14110425',"
         g_str_Parame = g_str_Parame & "                                    '14210425','14310425','14140423','14240423','14340423','14140424',"
         g_str_Parame = g_str_Parame & "                                    '14340424','14140425','14240425','14350423','14340425','14150423',"
         g_str_Parame = g_str_Parame & "                                    '14250423','14150424','14350424','14150425','14250425','14350425',"
         g_str_Parame = g_str_Parame & "                                    '14160423','14260423','14360423','14160424','14360424','14160425',"
         g_str_Parame = g_str_Parame & "                                    '14260425','14360425')"
         g_str_Parame = g_str_Parame & "     OR SUBSTR(HIPDET_CTACBL,1,10) IN ('1415041923','1425041923','1435041923','1415041924','1425041924','1435041924',"
         g_str_Parame = g_str_Parame & "                                       '1415041925','1425041925','1435041925','1416041923','1426041923','1436041923',"
         g_str_Parame = g_str_Parame & "                                       '1416041924','1426041924','1436041924','1416041925','1426041925','1436041925'))"
      End If
      
      g_str_Parame = g_str_Parame & "    AND HIPDET_CTACBL > 0"
      g_str_Parame = g_str_Parame & "    AND HIPDET_NUMITE > 0"
      g_str_Parame = g_str_Parame & "    AND (HIPDET_CLASIF = 'DEFICIENTE' OR HIPDET_CLASIF = 'DUDOSO' OR HIPDET_CLASIF = 'PERDIDA')"
      g_str_Parame = g_str_Parame & "  GROUP BY HIPDET_CODEMP,HIPDET_NOMEMP "
       
      'Ejecuta consulta
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
      End If
      
      If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then g_rst_Princi.MoveFirst
      
      Do Until g_rst_Princi.EOF
      
         If Not IsNull(g_rst_Princi!MTO_ACT_PESADO) Then
            r_dbl_Monto = g_rst_Princi!MTO_ACT_PESADO
            r_dbl_MtoSol = g_rst_Princi!MONTO_SOLES
            r_dbl_MtoDol = g_rst_Princi!MONTO_DOLARES
            r_str_IndTip = "HP"
            r_dbl_Total = 0
         End If
         If r_dbl_Monto > 0 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO RCC_CONMEN("
            g_str_Parame = g_str_Parame & "CONMEN_PERANO, "
            g_str_Parame = g_str_Parame & "CONMEN_PERMES, "
            g_str_Parame = g_str_Parame & "CONMEN_CODEMP, "
            g_str_Parame = g_str_Parame & "CONMEN_TIPREP, "
            g_str_Parame = g_str_Parame & "CONMEN_INDTIP, "
            g_str_Parame = g_str_Parame & "CONMEN_NOMEMP, "
            g_str_Parame = g_str_Parame & "CONMEN_MONTOT, "
            g_str_Parame = g_str_Parame & "CONMEN_NUMTOT, "
            g_str_Parame = g_str_Parame & "CONMEN_MTOSOL, "
            g_str_Parame = g_str_Parame & "CONMEN_MTODOL) "
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & "'" & Format(p_PerAno, "0000") & "', "
            g_str_Parame = g_str_Parame & "'" & Format(p_PerMes, "00") & "', "
            g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!HIPDET_CODEMP) & "' , "
            g_str_Parame = g_str_Parame & "'" & Trim(r_int_TipRep) & "' , "
            g_str_Parame = g_str_Parame & "'" & Trim(r_str_IndTip) & "' , "
            g_str_Parame = g_str_Parame & "'" & Trim(g_rst_Princi!HIPDET_NOMEMP) & "' , "
            g_str_Parame = g_str_Parame & r_dbl_Monto & ", "                                 'Format(Trim(r_dbl_Monto), "########0.00")
            g_str_Parame = g_str_Parame & IIf(r_dbl_Total = 0, "NULL", r_dbl_Total) & " , "  'Format(Trim(r_dbl_Total), "########0.00")
            g_str_Parame = g_str_Parame & r_dbl_MtoSol & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoDol & ")"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               Exit Sub
            End If
            r_dbl_Monto = 0
            r_dbl_Total = 0
            r_dbl_MtoSol = 0
            r_dbl_MtoDol = 0
         End If
         
         g_rst_Princi.MoveNext
      Loop
      DoEvents
      p_BarPro.FloodPercent = p_BarPro.FloodPercent + 15
   Next r_int_TipRep
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   DoEvents
   p_BarPro.FloodPercent = 100
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub dir_LisCar_Change()
   fil_LisArc.Path = dir_LisCar.Path
End Sub

Private Sub drv_LisUni_Change()
   dir_LisCar.Path = drv_LisUni.Drive
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(fil_LisArc)
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)

   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   Call fs_Carga_Empresas
   Call fs_Carga_Creditos
   Call fs_Carga_Clasif
   
   cmb_TipArc.Clear
   cmb_TipArc.AddItem "PLANO - RCC"
   cmb_TipArc.ItemData(cmb_TipArc.NewIndex) = 0
   cmb_TipArc.AddItem "EXCEL - ENTIDADES SUPERVISADAS"
   cmb_TipArc.ItemData(cmb_TipArc.NewIndex) = 1
End Sub

Private Sub fs_Limpia()
   Dim r_int_PerMes  As Integer
   Dim r_int_PerAno  As Integer

   If Month(date) = 1 Then
      r_int_PerMes = 12
      r_int_PerAno = Year(date) - 1
   Else
      r_int_PerMes = Month(date) - 1
      r_int_PerAno = Year(date)
   End If

   Call gs_BuscarCombo_Item(cmb_PerMes, r_int_PerMes)
   ipp_PerAno.Text = Format(r_int_PerAno, "0000")
   
   pnl_BarPro.FloodPercent = 0
   
'   If CInt(cmb_TipArc.ItemData(cmb_TipArc.ListIndex)) = 0 Then
'      fil_LisArc.Pattern = "rcc*.ope"
'   ElseIf CInt(cmb_TipArc.ItemData(cmb_TipArc.ListIndex)) = 1 Then
'      fil_LisArc.Pattern = "*.xls"
'   End If
End Sub
