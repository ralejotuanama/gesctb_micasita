VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_RegCom_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12360
   Icon            =   "GesCtb_frm_195.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel13 
      Height          =   8355
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12735
      _Version        =   65536
      _ExtentX        =   22463
      _ExtentY        =   14737
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
      Begin Threed.SSPanel frm_Ctb_AsiCtb_05 
         Height          =   615
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   12270
         _Version        =   65536
         _ExtentX        =   21643
         _ExtentY        =   1085
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
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   10980
            Top             =   60
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   555
            Left            =   690
            TabIndex        =   7
            Top             =   30
            Width           =   4755
            _Version        =   65536
            _ExtentX        =   8387
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Generación de Data para la Planilla"
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
            Left            =   30
            Picture         =   "GesCtb_frm_195.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   8
         Top             =   720
         Width           =   12270
         _Version        =   65536
         _ExtentX        =   21643
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
         Begin VB.CommandButton cmd_ExpTxt 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_195.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Generar Archivo de Texto"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11655
            Picture         =   "GesCtb_frm_195.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   885
         Left            =   60
         TabIndex        =   9
         Top             =   1410
         Width           =   12270
         _Version        =   65536
         _ExtentX        =   21643
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
         Begin VB.ComboBox cmb_TipCtb 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   150
            Width           =   2385
         End
         Begin VB.TextBox txt_NomArc 
            Height          =   315
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "txt_NomArc"
            Top             =   480
            Width           =   8835
         End
         Begin VB.CommandButton cmd_BuscaArc 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10890
            TabIndex        =   1
            ToolTipText     =   "Seleccionar archivo"
            Top             =   480
            Width           =   315
         End
         Begin VB.CommandButton cmd_Import 
            Height          =   585
            Left            =   11655
            Picture         =   "GesCtb_frm_195.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Importar archivo"
            Top             =   30
            Width           =   585
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo Contabilización:"
            Height          =   285
            Left            =   180
            TabIndex        =   14
            Top             =   150
            Width           =   1605
         End
         Begin VB.Label Label4 
            Caption         =   "Archivo a cargar:"
            Height          =   255
            Left            =   180
            TabIndex        =   11
            Top             =   510
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   5925
         Left            =   60
         TabIndex        =   12
         Top             =   2340
         Width           =   12270
         _Version        =   65536
         _ExtentX        =   21643
         _ExtentY        =   10451
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5740
            Left            =   90
            TabIndex        =   13
            Top             =   90
            Width           =   12090
            _ExtentX        =   21325
            _ExtentY        =   10134
            _Version        =   393216
            Rows            =   10
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_RegCom_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type arr_PlaEmp
   plaemp_Codigo    As String
   plaemp_Sueldo    As String
End Type
   
Dim l_arr_GenArc()      As arr_PlaEmp

Private Sub cmb_TipCtb_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_BuscaArc)
   End If
End Sub

Private Sub cmd_ExpTxt_Click()
   If cmb_TipCtb.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un tipo de contabilización.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCtb)
      Exit Sub
   End If
    
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
    
   Screen.MousePointer = 11
   Call fs_GenArchivo
   Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpiar
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(txt_NomArc)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipCtb, 1, "131")
   
   grd_Listad.ColWidth(0) = 1000 'Codigo
   grd_Listad.ColWidth(1) = 1300 'Tipo documento
   grd_Listad.ColWidth(2) = 1270 'Nro Documento
   grd_Listad.ColWidth(3) = 3500 'Nombres
   grd_Listad.ColWidth(4) = 1300 'Tipo Cuenta
   grd_Listad.ColWidth(5) = 2050 'Cuenta Corriente/INTERBANCARIA
   grd_Listad.ColWidth(6) = 1300 'Sueldo
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.Rows = 0
End Sub

Private Sub fs_Limpiar()
   Call gs_LimpiaGrid(grd_Listad)
   txt_NomArc.Text = ""
   Call gs_SetFocus(txt_NomArc)
   cmd_ExpTxt.Enabled = False
End Sub

Private Sub cmd_BuscaArc_Click()
   dlg_Guarda.Filter = "Archivos Excel |*.xlsx;*.xls"
   dlg_Guarda.ShowOpen
   txt_NomArc.Text = UCase(dlg_Guarda.FileName)
   Exit Sub
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub fs_GenArchivo()
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer
Dim r_int_NumRes    As Integer
Dim r_str_NomRes    As String
Dim r_str_Cadena    As String
Dim r_str_CadAux    As String
Dim r_dbl_PlaTot    As Double
Dim r_int_RegTot    As Integer
   
Dim r_arr_LogPro()  As modprc_g_tpo_LogPro
Dim r_int_NumIte    As Integer
Dim r_int_NumAsi    As Integer
Dim r_str_Glosa     As String
Dim r_dbl_MtoSol    As Double
Dim r_dbl_MtoDol    As Double
Dim r_str_FechaL    As String
Dim r_str_FechaC    As String
Dim r_int_NumLib    As Integer
Dim r_str_Origen    As String
Dim r_int_Contar    As Integer
Dim r_str_CtaHab    As String
Dim r_str_CtaDeb    As String
Dim r_dbl_TipSbs    As Double
Dim r_str_TipNot    As String
   
   r_int_PerAno = Year(moddat_g_str_FecSis)
   r_int_PerMes = Month(moddat_g_str_FecSis)
   r_str_NomRes = moddat_g_str_RutLoc & "\" & Format(moddat_g_str_FecSis, "yyyymm") & "_Haberes.TXT"
                      
   'Creando Archivo
   r_int_NumRes = FreeFile
   Open r_str_NomRes For Output As r_int_NumRes
      
   r_str_Cadena = ""
   r_dbl_PlaTot = 0
      
   r_int_RegTot = 0
   For r_int_Contar = 1 To grd_Listad.Rows - 1
       r_dbl_PlaTot = r_dbl_PlaTot + CDbl(grd_Listad.TextMatrix(r_int_Contar, 6))
       r_int_RegTot = r_int_RegTot + 1
   Next
   r_str_CadAux = ""
   For r_int_Contar = 1 To 68
       r_str_CadAux = r_str_CadAux & " "
   Next
   r_str_Cadena = r_str_Cadena & "70000110661000100040896PEN" & Format(r_dbl_PlaTot * 100, "000000000000000") & _
                                 "A" & Format(moddat_g_str_FecSis, "yyyymmdd") & "H" & "HABERES 5TA CATEGORIA    " & _
                                 Format(r_int_RegTot, "000000") & "S" & r_str_CadAux
   Print #r_int_NumRes, r_str_Cadena
      
   r_str_CadAux = ""
   For r_int_Contar = 1 To 101
       r_str_CadAux = r_str_CadAux & " "
   Next
   For r_int_Contar = 1 To grd_Listad.Rows - 1
       r_str_Cadena = ""
       r_str_Cadena = r_str_Cadena & "002" & Trim(grd_Listad.TextMatrix(r_int_Contar, 1)) & _
                                     Left(Trim(grd_Listad.TextMatrix(r_int_Contar, 2)) & "            ", 12) & _
                                     Trim(grd_Listad.TextMatrix(r_int_Contar, 4)) & _
                                     Trim(grd_Listad.TextMatrix(r_int_Contar, 5)) & _
                                     Left("HABERES" & Trim(grd_Listad.TextMatrix(r_int_Contar, 0)) & r_int_PerAno & Format(r_int_PerMes, "00") & "                                        ", 40) & _
                                     Format(grd_Listad.TextMatrix(r_int_Contar, 6) * 100, "000000000000000") & _
                                     Left("HABERES " & fs_nombresMes(r_int_PerMes) & "                                        ", 40) & _
                                     r_str_CadAux
       Print #r_int_NumRes, r_str_Cadena
   Next
   'Left(Trim(grd_Listad.TextMatrix(r_int_Contar, 3)) & "                                        ", 40) &
   
   'Cerrando Archivo Resumen
   Close r_int_NumRes
   '-----------------CONTABILIZAR PLANILLA------------------------------------
   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 12
   
   r_int_NumAsi = 0
   r_int_NumIte = 0
   r_str_FechaC = Format(moddat_g_str_FecSis, "yyyymmdd")
   r_str_FechaL = moddat_g_str_FecSis
                
   r_str_Glosa = Mid("", 1, 60)
   'Obteniendo Nro. de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
   r_str_CadAux = ""
   r_str_CadAux = Trim(cmb_TipCtb.Text)
   r_str_Glosa = "PLANILLA " & r_str_CadAux & "/" & r_int_PerAno & Format(r_int_PerMes, "00")
                 
   'Insertar en CABECERA
   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                 r_int_NumAsi, Format(1, "000"), r_dbl_TipSbs, r_str_TipNot, Trim(r_str_Glosa), r_str_FechaL, "1")
   If cmb_TipCtb.ItemData(cmb_TipCtb.ListIndex) = 1 Then 'HABERES
      r_str_CtaDeb = "251504010101"
      r_str_CtaHab = "111301060102"
   ElseIf cmb_TipCtb.ItemData(cmb_TipCtb.ListIndex) = 2 Then 'MOVILIDAD
      r_str_CtaDeb = "151702010101"
      r_str_CtaHab = "111301060102"
   ElseIf cmb_TipCtb.ItemData(cmb_TipCtb.ListIndex) = 3 Then 'CTS
      r_str_CtaDeb = "251509010103"
      r_str_CtaHab = "111301060102"
   ElseIf cmb_TipCtb.ItemData(cmb_TipCtb.ListIndex) = 4 Then 'GRATIFICACIONES
      r_str_CtaDeb = "251509010101"
      r_str_CtaHab = "111301060102"
   End If
   r_int_NumIte = 1
   For r_int_Contar = 1 To grd_Listad.Rows - 1
       r_dbl_MtoSol = CDbl(grd_Listad.TextMatrix(r_int_Contar, 6))
       r_dbl_MtoDol = 0
       r_str_Glosa = r_str_CadAux & " " & r_int_PerAno & Format(r_int_PerMes, "00") & "/" & _
                     Trim(grd_Listad.TextMatrix(r_int_Contar, 0)) & "/" & "HABERES  " & Format(r_int_PerMes, "00") & "-" & r_int_PerAno 'fs_ExtraeApellido(grd_Listad.TextMatrix(r_int_Contar, 3))
       Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                            r_int_NumAsi, r_int_NumIte, r_str_CtaDeb, CDate(r_str_FechaL), _
                                            r_str_Glosa, "D", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
       r_int_NumIte = r_int_NumIte + 1
   Next
   For r_int_Contar = 1 To grd_Listad.Rows - 1
       r_dbl_MtoSol = CDbl(grd_Listad.TextMatrix(r_int_Contar, 6))
       r_dbl_MtoDol = 0
       r_str_Glosa = r_str_CadAux & " " & r_int_PerAno & Format(r_int_PerMes, "00") & "/" & _
                     Trim(grd_Listad.TextMatrix(r_int_Contar, 0)) & "/" & "HABERES  " & Format(r_int_PerMes, "00") & "-" & r_int_PerAno 'fs_ExtraeApellido(grd_Listad.TextMatrix(r_int_Contar, 3))
       
       Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, _
                                            r_int_NumAsi, r_int_NumIte, r_str_CtaHab, CDate(r_str_FechaL), _
                                            r_str_Glosa, "H", r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FechaL))
       r_int_NumIte = r_int_NumIte + 1
   Next
   '-----------GUARDAR EN LOG-----------------------------------------
   r_str_CadAux = ""
   If cmb_TipCtb.ItemData(cmb_TipCtb.ListIndex) = 1 Then 'HABERES
      r_str_CadAux = "5"
   ElseIf cmb_TipCtb.ItemData(cmb_TipCtb.ListIndex) = 2 Then 'MOVILIDAD
      r_str_CadAux = "6"
   ElseIf cmb_TipCtb.ItemData(cmb_TipCtb.ListIndex) = 3 Then 'CTS
      r_str_CadAux = "7"
   ElseIf cmb_TipCtb.ItemData(cmb_TipCtb.ListIndex) = 4 Then 'GRATIFICACIONES
      r_str_CadAux = "8"
   End If
   Call fs_Actualiza_Proceso(r_int_PerAno, r_int_PerMes, CInt(r_str_CadAux), r_int_NumAsi)
   '-----------MENSAJE FINAL------------------------------------------
   MsgBox "Archivo generado con éxito: " & r_str_NomRes & vbCrLf & _
          "Asiento contable generado : " & r_int_NumAsi & vbCrLf & _
          "Nro de registros generados: " & r_int_NumIte - 1, vbInformation, modgen_g_str_NomPlt
End Sub

Function fs_ExtraeApellido(p_Apellido As String) As String
Dim r_int_PosIni As Integer
Dim r_int_PosFin As Integer

   r_int_PosIni = 0
   r_int_PosIni = InStr(1, Trim(p_Apellido), " ") + 1
   r_int_PosFin = InStr(r_int_PosIni, Trim(p_Apellido), " ")
   
   If r_int_PosFin = 0 And r_int_PosIni - 1 > 0 Then
      fs_ExtraeApellido = Trim(Mid(Trim(p_Apellido), 1, r_int_PosIni - 1))
   ElseIf r_int_PosFin > 0 Then
      fs_ExtraeApellido = Trim(Mid(Trim(p_Apellido), 1, r_int_PosFin))
   Else
      fs_ExtraeApellido = Trim(p_Apellido)
   End If
End Function

Function fs_nombresMes(p_Mes As Integer) As String
   Select Case p_Mes
          Case 1: fs_nombresMes = "ENERO"
          Case 2: fs_nombresMes = "FEBRERO"
          Case 3: fs_nombresMes = "MARZO"
          Case 4: fs_nombresMes = "ABRIL"
          Case 5: fs_nombresMes = "MAYO"
          Case 6: fs_nombresMes = "JUNIO"
          Case 7: fs_nombresMes = "JULIO"
          Case 8: fs_nombresMes = "AGOSTO"
          Case 9: fs_nombresMes = "SETIEMBRE"
          Case 10: fs_nombresMes = "OCTUBRE"
          Case 11: fs_nombresMes = "NOVIEMBRE"
          Case 12: fs_nombresMes = "DICIEMBRE"
   End Select
End Function

Public Sub fs_Actualiza_Proceso(ByVal p_PerAno As Integer, ByVal p_PerMes As Integer, ByVal p_TipPro As Integer, ByVal p_Asiento As Integer)
Dim r_str_Cadena     As String
Dim r_int_NumVec     As Integer
Dim r_rst_Record     As ADODB.Recordset

   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT COUNT(*) AS NUM_EJEC "
   r_str_Cadena = r_str_Cadena & "  FROM CTB_PERPRO "
   r_str_Cadena = r_str_Cadena & " WHERE PERPRO_CODANO = " & CStr(p_PerAno) & " "
   r_str_Cadena = r_str_Cadena & "   AND PERPRO_CODMES = " & CStr(p_PerMes) & " "
   r_str_Cadena = r_str_Cadena & "   AND PERPRO_TIPPRO = " & CStr(p_TipPro) & " "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Record, 3) Then
      Exit Sub
   End If
   
   r_rst_Record.MoveFirst
   r_int_NumVec = r_rst_Record!NUM_EJEC
   
   r_rst_Record.Close
   Set r_rst_Record = Nothing
   
   If r_int_NumVec = 0 Then
      'Inserta  PERPRO_RUTFIL VARCHAR2(200), PERPRO_NUMASI NUMBER(10)
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & "INSERT INTO CTB_PERPRO ("
      r_str_Cadena = r_str_Cadena & " PERPRO_CODANO, PERPRO_CODMES, PERPRO_TIPPRO, PERPRO_FECPRO, PERPRO_INDEJE, "
      r_str_Cadena = r_str_Cadena & " PERPRO_RUTFIL, PERPRO_NUMASI, "
      r_str_Cadena = r_str_Cadena & " SEGUSUCRE, SEGFECCRE, SEGHORCRE, SEGPLTCRE, SEGTERCRE, SEGSUCCRE) "
      r_str_Cadena = r_str_Cadena & "VALUES("
      r_str_Cadena = r_str_Cadena & " " & CStr(p_PerAno) & ", "
      r_str_Cadena = r_str_Cadena & " " & CStr(p_PerMes) & ", "
      r_str_Cadena = r_str_Cadena & " " & CStr(p_TipPro) & ", "
      r_str_Cadena = r_str_Cadena & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
      r_str_Cadena = r_str_Cadena & " " & 1 & ", "
      r_str_Cadena = r_str_Cadena & "'" & Trim(txt_NomArc.Text) & "', "
      r_str_Cadena = r_str_Cadena & "'" & CStr(p_Asiento) & "', "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "
      r_str_Cadena = r_str_Cadena & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
      r_str_Cadena = r_str_Cadena & " " & Format(Time, "HHMMSS") & ", "
      r_str_Cadena = r_str_Cadena & "'" & UCase(App.EXEName) & "', "
      r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "
      r_str_Cadena = r_str_Cadena & "'001'" & ") "
      
      If Not gf_EjecutaSQL(r_str_Cadena, modprc_g_rst_Grabar, 2) Then
         Exit Sub
      End If
      
   Else
      'Actualiza
      r_int_NumVec = r_int_NumVec + 1
      
      r_str_Cadena = ""
      r_str_Cadena = r_str_Cadena & "UPDATE CTB_PERPRO "
      r_str_Cadena = r_str_Cadena & "   SET PERPRO_INDEJE = PERPRO_INDEJE + 1, "
      r_str_Cadena = r_str_Cadena & "       PERPRO_RUTFIL = '" & Trim(txt_NomArc.Text) & "', "
      r_str_Cadena = r_str_Cadena & "       PERPRO_NUMASI =  TRIM(PERPRO_NUMASI) || '-" & CStr(p_Asiento) & "', "
      r_str_Cadena = r_str_Cadena & "       SEGUSUACT     = '" & Trim(modgen_g_str_CodUsu) & "', "
      r_str_Cadena = r_str_Cadena & "       SEGFECACT     = " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
      r_str_Cadena = r_str_Cadena & "       SEGHORACT     = " & Format(Time, "HHMMSS") & ", "
      r_str_Cadena = r_str_Cadena & "       SEGPLTACT     = '" & Trim(UCase(App.EXEName)) & "', "
      r_str_Cadena = r_str_Cadena & "       SEGTERACT     = '" & Trim(modgen_g_str_NombPC) & "', "
      r_str_Cadena = r_str_Cadena & "       SEGSUCACT     = '001' "
      r_str_Cadena = r_str_Cadena & " WHERE PERPRO_CODANO = " & CStr(p_PerAno) & " "
      r_str_Cadena = r_str_Cadena & "   AND PERPRO_CODMES = " & CStr(p_PerMes) & " "
      r_str_Cadena = r_str_Cadena & "   AND PERPRO_TIPPRO = " & CStr(p_TipPro) & " "
      
      If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Record, 2) Then
         Exit Sub
      End If
   
   End If
End Sub

Private Sub cmd_Import_Click()
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_int_Contad        As Integer
Dim r_str_CadAux        As String
Dim r_bol_Estado        As Boolean
Dim r_int_FilErr        As Integer
Dim r_int_FilCar        As Integer
Dim r_str_CodErr        As String
Dim r_int_Contar        As Integer

   If Len(Trim(txt_NomArc.Text)) = 0 Then
      MsgBox "Debe ingresar la ubicación y nombre del archivo a importar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If

   Call gs_LimpiaGrid(grd_Listad)
   r_int_FilErr = 0
   r_str_CodErr = ""
   r_int_FilCar = 0
   Screen.MousePointer = 11
   '-----------------------
   'Cabecera de la Grilla
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.FixedRows = 1
   grd_Listad.Rows = grd_Listad.Rows - 1
   grd_Listad.Row = 0
            
   grd_Listad.Col = 0:   grd_Listad.Text = "Código":                grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 1:   grd_Listad.Text = "Tipo Documento":        grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 2:   grd_Listad.Text = "Nro Documento":         grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 3:   grd_Listad.Text = "Apellidos y Nombres":   grd_Listad.CellAlignment = flexAlignLeftCenter
   grd_Listad.Col = 4:   grd_Listad.Text = "Tipo Cuenta":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 5:   grd_Listad.Text = "Cuenta Corriente":      grd_Listad.CellAlignment = flexAlignLeftCenter
   grd_Listad.Col = 6:   grd_Listad.Text = "Sueldo":                grd_Listad.CellAlignment = flexAlignRightCenter
   Call gs_RefrescaGrid(grd_Listad)
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=txt_NomArc.Text
   r_int_FilExc = 2
   r_str_CadAux = ""
   ReDim l_arr_GenArc(0)
   
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
      ReDim Preserve l_arr_GenArc(UBound(l_arr_GenArc) + 1)
      l_arr_GenArc(UBound(l_arr_GenArc)).plaemp_Codigo = Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value & "")
      l_arr_GenArc(UBound(l_arr_GenArc)).plaemp_Sueldo = Trim(r_obj_Excel.Cells(r_int_FilExc, 2).Value & "")
      
      r_str_CadAux = r_str_CadAux & "'" & Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value & " ") & "',"
      
      r_int_FilExc = r_int_FilExc + 1
   Loop
   If Len(Trim(r_str_CadAux)) > 0 Then
      r_str_CadAux = Mid(Trim(r_str_CadAux), 1, Len(Trim(r_str_CadAux)) - 1)
   End If
   
   g_str_Parame = ""
   'g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_CODSIC, DECODE(A.MAEPRV_TIPDOC,1,'L',4,'E') AS TIPDOC, A.MAEPRV_NUMDOC,  "
   'DECODE(A.MAEPRV_TIPDOC,1,'L',4,'E') AS TIPDOC,
   
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_CODSIC, A.MAEPRV_NUMDOC, CASE A.MAEPRV_TIPDOC "
   g_str_Parame = g_str_Parame & "                                            WHEN 1 THEN 'L' "
   g_str_Parame = g_str_Parame & "                                            WHEN 4 THEN 'E' "
   g_str_Parame = g_str_Parame & "                                            WHEN 7 THEN 'P' "
   g_str_Parame = g_str_Parame & "                                          END AS TIPDOC, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_RAZSOC, DECODE(A.MAEPRV_CODBNC_MN1,11,'P','I') AS TIPCTA,  "
   g_str_Parame = g_str_Parame & "        DECODE(A.MAEPRV_CODBNC_MN1,11,SUBSTR(TRIM(A.MAEPRV_CTACRR_MN1),1,8) || '00'|| SUBSTR(TRIM(A.MAEPRV_CTACRR_MN1),9,10),  "
   g_str_Parame = g_str_Parame & "               A.MAEPRV_NROCCI_MN1) AS NUM_CUENTA  "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A  "
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPPER = 2  "
   g_str_Parame = g_str_Parame & "    AND A.MAEPRV_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "    AND TRIM(A.MAEPRV_CODSIC) IN  (" & r_str_CadAux & ")  "
   g_str_Parame = g_str_Parame & "  ORDER BY A.MAEPRV_RAZSOC ASC  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Sub
   End If
      
   For r_int_Contar = 1 To UBound(l_arr_GenArc)
       grd_Listad.Rows = grd_Listad.Rows + 1
       grd_Listad.Row = grd_Listad.Rows - 1
       
       grd_Listad.Col = 0
       grd_Listad.Text = l_arr_GenArc(r_int_Contar).plaemp_Codigo
       r_bol_Estado = False
       
       g_rst_Princi.MoveFirst
       Do While Not g_rst_Princi.EOF
          If Trim(l_arr_GenArc(r_int_Contar).plaemp_Codigo) = Trim(g_rst_Princi!MAEPRV_CODSIC & "") Then
             r_bol_Estado = True
             grd_Listad.Col = 1
             grd_Listad.Text = CStr(g_rst_Princi!TIPDOC & "")
      
             grd_Listad.Col = 2
             grd_Listad.Text = CStr(g_rst_Princi!MAEPRV_NUMDOC & "")
      
             grd_Listad.Col = 3
             grd_Listad.Text = CStr(g_rst_Princi!MaePrv_RazSoc & "")
            
             grd_Listad.Col = 4
             grd_Listad.Text = CStr(g_rst_Princi!TIPCTA & "")
      
             grd_Listad.Col = 5
             grd_Listad.Text = CStr(g_rst_Princi!NUM_CUENTA & "")
             Exit Do
          End If
          g_rst_Princi.MoveNext
       Loop
       grd_Listad.Col = 6
       grd_Listad.Text = Format(l_arr_GenArc(r_int_Contar).plaemp_Sueldo, "###,###,##0.00")
       
       If r_bol_Estado = False Then
          r_int_FilErr = r_int_FilErr + 1
          r_str_CodErr = r_str_CodErr & " -" & l_arr_GenArc(r_int_Contar).plaemp_Codigo
       Else
          r_int_FilCar = r_int_FilCar + 1
       End If
   Next
   If r_str_CodErr = "" Then
      r_str_CodErr = "0"
   End If
   If grd_Listad.Rows > 0 Then
      MsgBox "Nro.de reg. observados : " & r_int_FilErr & vbCrLf & _
             "Código reg. observados : " & r_str_CodErr & vbCrLf & _
             "Nro.de reg. sin observar: " & r_int_FilCar & vbCrLf & _
             "Total  de  reg. cargados : " & grd_Listad.Rows - 1, vbInformation, modgen_g_str_NomPlt
   End If
   If r_int_FilCar = grd_Listad.Rows - 1 Then
      cmd_ExpTxt.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Quit
   Set r_obj_Excel = Nothing
   '-----------------------
   Screen.MousePointer = 0
End Sub

