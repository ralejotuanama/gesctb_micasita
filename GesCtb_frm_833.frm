VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_Provis_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   3135
   ClientLeft      =   5010
   ClientTop       =   4710
   ClientWidth     =   10425
   Icon            =   "GesCtb_frm_833.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10485
      _Version        =   65536
      _ExtentX        =   18494
      _ExtentY        =   5900
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
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   750
         Width           =   10335
         _Version        =   65536
         _ExtentX        =   18230
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_833.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9660
            Picture         =   "GesCtb_frm_833.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   10335
         _Version        =   65536
         _ExtentX        =   18230
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
            Height          =   495
            Left            =   630
            TabIndex        =   6
            Top             =   60
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de Provisiones"
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
            Picture         =   "GesCtb_frm_833.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1185
         Left            =   30
         TabIndex        =   7
         Top             =   1920
         Width           =   10335
         _Version        =   65536
         _ExtentX        =   18230
         _ExtentY        =   2090
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
         Begin VB.ComboBox cmb_ClaPrv 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   2805
         End
         Begin Threed.SSPanel pnl_PrvGen 
            Height          =   315
            Left            =   2280
            TabIndex        =   17
            Top             =   390
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
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
         Begin Threed.SSPanel pnl_PrvEsp 
            Height          =   315
            Left            =   2280
            TabIndex        =   18
            Top             =   720
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
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
         Begin Threed.SSPanel pnl_PrvCam 
            Height          =   315
            Left            =   7470
            TabIndex        =   19
            Top             =   390
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
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
         Begin Threed.SSPanel pnl_PrvCic 
            Height          =   315
            Left            =   7470
            TabIndex        =   21
            Top             =   720
            Width           =   2775
            _Version        =   65536
            _ExtentX        =   4895
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
         Begin VB.Label Label6 
            Caption         =   "Prov. ProCiclica:"
            Height          =   285
            Left            =   5340
            TabIndex        =   20
            Top             =   780
            Width           =   1425
         End
         Begin VB.Label Label5 
            Caption         =   "Prov. Ries. Camb.:"
            Height          =   285
            Left            =   5340
            TabIndex        =   11
            Top             =   420
            Width           =   1425
         End
         Begin VB.Label Label4 
            Caption         =   "Prov. Específica:"
            Height          =   285
            Left            =   150
            TabIndex        =   10
            Top             =   750
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Prov. Genérica:"
            Height          =   285
            Left            =   150
            TabIndex        =   9
            Top             =   420
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Clasificación:"
            Height          =   315
            Left            =   150
            TabIndex        =   8
            Top             =   120
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   435
         Left            =   30
         TabIndex        =   12
         Top             =   1440
         Width           =   10335
         _Version        =   65536
         _ExtentX        =   18230
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1320
            TabIndex        =   13
            Top             =   60
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
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
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   4110
            TabIndex        =   14
            Top             =   60
            Width           =   6135
            _Version        =   65536
            _ExtentX        =   10821
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
         Begin VB.Label Label2 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   3450
            TabIndex        =   15
            Top             =   60
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_Provis_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Grabar_Click()

   If MsgBox("¿Está seguro que desea registrar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'g_str_Parame = "UPDATE CRE_HIPCIE SET "
   'g_str_Parame = g_str_Parame & "HIPCIE_CLAPRV = " & cmb_ClaPrv.ListIndex & " "
   'g_str_Parame = g_str_Parame & "WHERE HIPCIE_TDOCLI = '" & moddat_g_int_TipDoc & "' AND "
   'g_str_Parame = g_str_Parame & "HIPCIE_NDOCLI = " & moddat_g_str_NumDoc & " AND "
   'g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & Right(modsec_g_str_Period, 2) & " AND "
   'g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & Left(modsec_g_str_Period, 4) & " "
   
   'If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
   '    Exit Sub
   'End If
   
   Call fs_CalPrv
   
   Unload Me

End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call gs_CentraForm(Me)
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call modsec_gs_Carga_ClaPrv(cmb_ClaPrv)
   Call cmd_Buscar_Click
   
   pnl_NumOpe.Caption = Left(moddat_g_str_NumOpe, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Right(moddat_g_str_NumOpe, 5)
   pnl_Client.Caption = moddat_g_int_TipDoc & "-" & moddat_g_str_NumDoc & " / " & moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Buscar_Click()

   g_str_Parame = "SELECT * FROM CRE_HIPCIE WHERE "
   g_str_Parame = g_str_Parame & "HIPCIE_TDOCLI = " & moddat_g_int_TipDoc & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_NDOCLI = '" & moddat_g_str_NumDoc & "' AND "
   'g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & Right("201009", 2) & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & Right(modsec_g_str_Period, 2) & " AND "
   g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & Left(modsec_g_str_Period, 4) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   pnl_PrvGen.Caption = Format(g_rst_Princi!HIPCIE_PRVGEN, "###,###,##0.00") & " "
   pnl_PrvEsp.Caption = Format(g_rst_Princi!HIPCIE_PRVESP, "###,###,##0.00") & " "
   pnl_PrvCam.Caption = Format(g_rst_Princi!HIPCIE_PRVCAM, "###,###,##0.00") & " "
   pnl_PrvCic.Caption = Format(g_rst_Princi!HIPCIE_PRVCIC, "###,###,##0.00") & " "
        
   cmb_ClaPrv.ListIndex = g_rst_Princi!HIPCIE_CLAPRV
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
End Sub

Public Sub fs_CalPrv()
   'Código Proceso   :  CTBP1009
   'Descripción      :  Cálculo de Provisiones
   'Resumen          :  Calculo de Provisiones
   'F. Creación      :  31-07-2009
   'U. Creación      :  Miguel Angel Ikehara Punk
   'F. Actualización :
   'U. Actualización :

   Dim r_lng_NumReg        As Long
   Dim r_lng_TotReg        As Long
   Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
   
   Dim r_str_FecPro        As String
   
   Dim r_rst_Grabar        As ADODB.Recordset
   
   Dim r_dbl_TipCam_Dol    As Double
   
   Dim r_arr_DetGar()      As modprc_g_tpo_DetGar
   Dim r_arr_TipPrv()      As modprc_g_tpo_TipPrv
   
   Dim r_int_ClaGar        As Integer
   Dim r_int_Contad        As Integer
   
   Dim r_dbl_PrvGen        As Double
   Dim r_dbl_PrvEsp        As Double
   Dim r_dbl_PrvCic        As Double
   Dim r_dbl_PrvCam        As Double
   
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1009"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   'Fecha de Proceso
   r_str_FecPro = Format(date, "dd/mm/yyyy")
   
   'Obteniendo Tipo de Cambio de Cierre
   'r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(p_FecFin), "yyyymmdd"), 2)
   'r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Left(modsec_g_str_Period, 4) & Format(Right(modsec_g_str_Period, 2), "00") & "01", 2)
   
   r_dbl_TipCam_Dol = moddat_gf_ObtieneTipCamDia(2, 2, Left(modsec_g_str_Period, 4) & Format(Right(modsec_g_str_Period, 2), "00") & Format(modsec_gf_Fin_Del_Mes("01/" & Format(Right(modsec_g_str_Period, 2), "00") & "/" & Left(modsec_g_str_Period, 4)), "dd"), 2)
   
   'Leer Tablas de Provisiones para Créditos Hipotecarios
   modprc_g_str_CadEje = "SELECT * FROM CTB_TIPPRV WHERE TIPPRV_CLACRE = '13' "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      'r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      
      'Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, _
                                       "Error al Leer Tabla CTB_TIPPRV.")
      
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
      'r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      
      'Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, _
      '                                 "Error al Leer Tabla CTB_DETGAR.")
      
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
      
   'Leyendo Cursor Principal (Créditos Hipotecarios)
   modprc_g_str_CadEje = "SELECT * FROM CRE_HIPCIE WHERE HIPCIE_PERMES = " & Right(modsec_g_str_Period, 2) & " AND HIPCIE_PERANO = " & Left(modsec_g_str_Period, 4) & " AND "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "HIPCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Princi, 3) Then
      'r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      
      'Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, _
                                       "Error al Leer Tabla CRE_HIPCIE.")
      
      modprc_g_rst_Princi.Close
      Set modprc_g_rst_Princi = Nothing
                        
      Exit Sub
   End If

   If Not (modprc_g_rst_Princi.BOF And modprc_g_rst_Princi.EOF) Then
      modprc_g_rst_Princi.MoveFirst
      
      Do While Not modprc_g_rst_Princi.EOF
         'Determinando Clase de Garantía según Garantía
         r_int_ClaGar = 0
         For r_int_Contad = 1 To UBound(r_arr_DetGar)
            If r_arr_DetGar(r_int_Contad).DetGar_Codigo = modprc_g_rst_Princi!HIPCIE_TIPGAR Then
               r_int_ClaGar = r_arr_DetGar(r_int_Contad).DetGar_ClaGar
               Exit For
            End If
         Next r_int_Contad
         
         r_dbl_PrvGen = 0
         r_dbl_PrvCic = 0
         r_dbl_PrvCam = 0
         r_dbl_PrvEsp = 0
         
         If CStr(cmb_ClaPrv.ListIndex) = 0 Then
            'Calculando Provisión Generica
            If modprc_g_rst_Princi!HIPCIE_CODPRD = "001" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "003" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "004" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "007" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "009" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "010" Then
               If CDate(gf_FormatoFecha(modprc_g_rst_Princi!HIPCIE_FECDES)) <= CDate("30/06/2010") Then
                  r_dbl_PrvGen = modprc_gf_PorcenProv(r_arr_TipPrv, 1, CStr(cmb_ClaPrv.ListIndex), r_int_ClaGar) / 100 * (((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON) - modprc_g_rst_Princi!HIPCIE_INTDIF) * 2 / 3)
               Else
                  r_dbl_PrvGen = modprc_gf_PorcenProv(r_arr_TipPrv, 1, CStr(cmb_ClaPrv.ListIndex), r_int_ClaGar) / 100 * (((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON) - modprc_g_rst_Princi!HIPCIE_INTDIF))
               End If
            Else
               r_dbl_PrvGen = modprc_gf_PorcenProv(r_arr_TipPrv, 1, CStr(cmb_ClaPrv.ListIndex), r_int_ClaGar) / 100 * (((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON) - modprc_g_rst_Princi!HIPCIE_INTDIF))
            End If
                        
             'Calculando Provisión Pro-Ciclica
            If modprc_g_rst_Princi!HIPCIE_CODPRD = "001" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "003" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "004" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "007" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "009" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "010" Then
               r_dbl_PrvCic = modprc_gf_PorcenProv(r_arr_TipPrv, 3, CStr(cmb_ClaPrv.ListIndex), r_int_ClaGar) / 100 * (((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON) - modprc_g_rst_Princi!HIPCIE_INTDIF) * 2 / 3)
            Else
               r_dbl_PrvCic = modprc_gf_PorcenProv(r_arr_TipPrv, 3, CStr(cmb_ClaPrv.ListIndex), r_int_ClaGar) / 100 * ((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON) - modprc_g_rst_Princi!HIPCIE_INTDIF)
            End If
                        
            If modprc_g_rst_Princi!HIPCIE_TIPMON = 2 Then
               'Calculando Provisión Riesgo Cambiario
               r_dbl_PrvCam = modprc_gf_PorcenProv(r_arr_TipPrv, 4, CStr(cmb_ClaPrv.ListIndex), r_int_ClaGar) / 100 * (modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON)
            End If
         Else
            
            'Calculando Provisión Específica
            If modprc_g_rst_Princi!HIPCIE_CODPRD = "001" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "003" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "004" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "007" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "009" Or modprc_g_rst_Princi!HIPCIE_CODPRD = "010" Then
               If CDate(gf_FormatoFecha(modprc_g_rst_Princi!HIPCIE_FECDES)) <= CDate("30/06/2010") Then
                  r_dbl_PrvEsp = modprc_gf_PorcenProv(r_arr_TipPrv, 2, CStr(cmb_ClaPrv.ListIndex), r_int_ClaGar) / 100 * (((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON) - modprc_g_rst_Princi!HIPCIE_INTDIF) * 2 / 3)
               Else
                  r_dbl_PrvEsp = modprc_gf_PorcenProv(r_arr_TipPrv, 2, CStr(cmb_ClaPrv.ListIndex), r_int_ClaGar) / 100 * (((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON) - modprc_g_rst_Princi!HIPCIE_INTDIF))
               End If
            Else
               r_dbl_PrvEsp = modprc_gf_PorcenProv(r_arr_TipPrv, 2, CStr(cmb_ClaPrv.ListIndex), r_int_ClaGar) / 100 * (((modprc_g_rst_Princi!HIPCIE_SALCAP + modprc_g_rst_Princi!HIPCIE_SALCON) - modprc_g_rst_Princi!HIPCIE_INTDIF))
            End If
         End If
            
         'If modprc_g_rst_Princi!HIPCIE_TIPMON = 2 Then
         '   r_dbl_PrvGen = r_dbl_PrvGen * r_dbl_TipCam_Dol
         '   r_dbl_PrvEsp = r_dbl_PrvEsp * r_dbl_TipCam_Dol
         '   r_dbl_PrvCic = r_dbl_PrvCic * r_dbl_TipCam_Dol
         '   r_dbl_PrvCam = r_dbl_PrvCam * r_dbl_TipCam_Dol
         'End If
            
         r_dbl_PrvGen = CDbl(Format(r_dbl_PrvGen, "######0.00"))
         r_dbl_PrvEsp = CDbl(Format(r_dbl_PrvEsp, "######0.00"))
         r_dbl_PrvCic = CDbl(Format(r_dbl_PrvCic, "######0.00"))
         r_dbl_PrvCam = CDbl(Format(r_dbl_PrvCam, "######0.00"))
            
         
         'g_str_Parame = "UPDATE CRE_HIPCIE SET "
         'g_str_Parame = g_str_Parame & "HIPCIE_CLAPRV = " & cmb_ClaPrv.ListIndex & " "
         'g_str_Parame = g_str_Parame & "WHERE HIPCIE_TDOCLI = '" & moddat_g_int_TipDoc & "' AND "
         'g_str_Parame = g_str_Parame & "HIPCIE_NDOCLI = " & moddat_g_str_NumDoc & " AND "
         'g_str_Parame = g_str_Parame & "HIPCIE_PERMES = " & Right(modsec_g_str_Period, 2) & " AND "
         'g_str_Parame = g_str_Parame & "HIPCIE_PERANO = " & Left(modsec_g_str_Period, 4) & " "
   
         
         
         
         'Actualizando Provisión en CRE_HIPCIE
         modprc_g_str_CadEje = "UPDATE CRE_HIPCIE SET " & _
                               "HIPCIE_CLAPRV = " & CStr(cmb_ClaPrv.ListIndex) & ", " & _
                               "HIPCIE_PRVGEN = " & CStr(r_dbl_PrvGen) & ", " & _
                               "HIPCIE_PRVESP = " & CStr(r_dbl_PrvEsp) & ", " & _
                               "HIPCIE_PRVCIC = " & CStr(r_dbl_PrvCic) & ", " & _
                               "HIPCIE_PRVCAM = " & CStr(r_dbl_PrvCam) & " WHERE " & _
                               "HIPCIE_NUMOPE = '" & modprc_g_rst_Princi!HIPCIE_NUMOPE & "' AND HIPCIE_PERMES = " & Right(modsec_g_str_Period, 2) & " AND HIPCIE_PERANO = " & Left(modsec_g_str_Period, 4) & " "
         
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, r_rst_Grabar, 2) Then
            'r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
            
            'Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, _
                                             "Error al actualizar Tabla CRE_HIPCIE - Operación: " & modprc_g_rst_Princi!HIPCIE_NUMOPE)
         End If
         
            
         'Leyendo siguiente registro
         modprc_g_rst_Princi.MoveNext
             
         r_lng_NumReg = r_lng_NumReg + 1
            
         'p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))

         DoEvents
      Loop
      
      'p_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
   'Else
      'r_arr_LogPro(1).LogPro_NumErr = r_arr_LogPro(1).LogPro_NumErr + 1
      
      'Call modprc_gs_GrabaErrorProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, r_arr_LogPro(1).LogPro_NumErr, _
                                       "No existen registros en tabla CRE_HIPCIE.")
   End If
   
   modprc_g_rst_Princi.Close
   Set modprc_g_rst_Princi = Nothing
   
   
   'Call modprc_gs_GrabaCabeceraLogProceso(r_arr_LogPro(1).LogPro_CodPro, r_arr_LogPro(1).LogPro_FInEje, r_arr_LogPro(1).LogPro_HInEje, _
                                          Format(Date, "yyyymmdd"), Format(Time, "hhmmss"), r_lng_NumReg, r_arr_LogPro(1).LogPro_NumErr, _
                                          p_CodEmp, "", 0, p_PerMes, p_PerAno, "0", "0")
End Sub


