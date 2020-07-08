VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_MtoItf_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   5040
   ClientLeft      =   3705
   ClientTop       =   3150
   ClientWidth     =   9030
   Icon            =   "GesCtb_frm_012.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9105
      _Version        =   65536
      _ExtentX        =   16060
      _ExtentY        =   9340
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
         TabIndex        =   1
         Top             =   30
         Width           =   8955
         _Version        =   65536
         _ExtentX        =   15796
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
            TabIndex        =   2
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Mantenimiento de ITF"
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
            Picture         =   "GesCtb_frm_012.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1155
         Left            =   60
         TabIndex        =   3
         Top             =   1440
         Width           =   8925
         _Version        =   65536
         _ExtentX        =   15743
         _ExtentY        =   2037
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1890
            MaxLength       =   12
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   7
            Top             =   720
            Width           =   1065
         End
         Begin VB.Label Label8 
            Caption         =   "Búsqueda por Documento de Identidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   6
            Top             =   60
            Width           =   3885
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2400
         Left            =   30
         TabIndex        =   9
         Top             =   2610
         Width           =   8955
         _Version        =   65536
         _ExtentX        =   15796
         _ExtentY        =   4233
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
         BorderWidth     =   1
         BevelOuter      =   0
         BevelInner      =   1
         Begin MSFlexGridLib.MSFlexGrid grd_LisOpe 
            Height          =   1965
            Left            =   60
            TabIndex        =   10
            Top             =   390
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   3466
            _Version        =   393216
            Rows            =   21
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   90
            TabIndex        =   11
            Top             =   90
            Width           =   700
            _Version        =   65536
            _ExtentX        =   1235
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Periodo"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   780
            TabIndex        =   12
            Top             =   90
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Declarante"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   6570
            TabIndex        =   13
            Top             =   90
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1746
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mto. Soles"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   285
            Left            =   2280
            TabIndex        =   14
            Top             =   90
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fec. Mov."
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   3270
            TabIndex        =   15
            Top             =   90
            Width           =   3300
            _Version        =   65536
            _ExtentX        =   5821
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Movimiento"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   7560
            TabIndex        =   16
            Top             =   90
            Width           =   990
            _Version        =   65536
            _ExtentX        =   1746
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ITF Soles"
            ForeColor       =   16777215
            BackColor       =   16384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   645
         Left            =   30
         TabIndex        =   17
         Top             =   750
         Width           =   8955
         _Version        =   65536
         _ExtentX        =   15796
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
         Begin VB.CommandButton cmd_BusOpe 
            Height          =   585
            Left            =   1830
            Picture         =   "GesCtb_frm_012.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Buscar Crédito por Número de Operación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_012.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Modificar Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_012.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Nueva Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "GesCtb_frm_012.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Borrar Ficha"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   2430
            Picture         =   "GesCtb_frm_012.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   8340
            Picture         =   "GesCtb_frm_012.frx":1248
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Mnt_MtoItf_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:     txt_NumDoc.MaxLength = 8
         Case Else:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   frm_Mnt_MtoItf_01.Show 1
End Sub

Private Sub cmd_Borrar_Click()
   Dim r_str_PerMes As String
   Dim r_str_PerAno As String
   
   If grd_LisOpe.Rows = 0 Then
      Exit Sub
   End If

   If MsgBox("¿Está seguro que desea borrar el registro?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   grd_LisOpe.Col = 6
   
   If Trim(grd_LisOpe.Text) = 2 Then
      Call gs_RefrescaGrid(grd_LisOpe)
      MsgBox "Solo se pueden eliminar los ingresos manuales.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   grd_LisOpe.Col = 0
   r_str_PerMes = Mid(Trim(grd_LisOpe.Text), 6, 2)
   r_str_PerAno = Mid(Trim(grd_LisOpe.Text), 1, 4)
     
   'Instrucción SQL
   g_str_Parame = "DELETE FROM CTB_DETITF WHERE DETITF_PERMES = " & r_str_PerMes & " AND DETITF_PERANO = " & r_str_PerAno & " AND "
   g_str_Parame = g_str_Parame & "DETITF_TIPDOC = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "DETITF_NUMDOC = " & Trim(txt_NumDoc.Text) & " "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Screen.MousePointer = 0
      Exit Sub
       
   End If
   
   MsgBox "Registro eliminado.", vbInformation, modgen_g_str_NomPlt
   
   Call gs_RefrescaGrid(grd_LisOpe)
   Call cmd_BusOpe_Click
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Editar_Click()
   If grd_LisOpe.Rows = 0 Then
      Exit Sub
   End If
   
   grd_LisOpe.Col = 6
   
   If Trim(grd_LisOpe.Text) = 2 Then
      Call gs_RefrescaGrid(grd_LisOpe)
      MsgBox "Solo se pueden modificar los ingresos manuales.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   grd_LisOpe.Col = 0
   modsec_g_str_Period = Trim(grd_LisOpe.Text)
   moddat_g_str_TipDoc = CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex))
   moddat_g_str_NumDoc = Trim(txt_NumDoc.Text)
   
   grd_LisOpe.Col = 4
   modsec_g_dbl_MtoSol = Trim(grd_LisOpe.Text)
   
   grd_LisOpe.Col = 5
   modsec_g_dbl_ITFSol = Trim(grd_LisOpe.Text)
   
   grd_LisOpe.Col = 7
   moddat_g_int_TipCli = Trim(grd_LisOpe.Text)
   
   Call gs_RefrescaGrid(grd_LisOpe)
   moddat_g_int_FlgGrb = 2
   frm_Mnt_MtoItf_01.Show 1
   Call cmd_BusOpe_Click
   
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   Call cmd_Limpia_Click
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call gs_CentraForm(Me)
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicio
   Call cmd_Limpia_Click
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   
   grd_LisOpe.ColWidth(0) = 700
   grd_LisOpe.ColWidth(1) = 1450
   grd_LisOpe.ColWidth(2) = 1000
   grd_LisOpe.ColWidth(3) = 3300
   grd_LisOpe.ColWidth(4) = 990
   grd_LisOpe.ColWidth(5) = 990
   grd_LisOpe.ColWidth(6) = 0
   grd_LisOpe.ColWidth(7) = 0
   
   grd_LisOpe.ColAlignment(0) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(1) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(2) = flexAlignCenterCenter
   grd_LisOpe.ColAlignment(3) = flexAlignLeftCenter
   grd_LisOpe.ColAlignment(4) = flexAlignRightCenter
   grd_LisOpe.ColAlignment(5) = flexAlignRightCenter
   grd_LisOpe.ColAlignment(6) = flexAlignRightCenter
   grd_LisOpe.ColAlignment(7) = flexAlignRightCenter

   Call gs_LimpiaGrid(grd_LisOpe)
   
   cmb_TipDoc.AddItem "DNI"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(1)
   
   cmb_TipDoc.AddItem "CARNE DE EXTRANJERIA"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(2)
   
   cmb_TipDoc.AddItem "PASAPORTE"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(5)
   
   cmb_TipDoc.AddItem "RUC"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = CInt(6)
         
   'Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   
End Sub

Private Sub cmd_Limpia_Click()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
   Call gs_SetFocus(cmb_TipDoc)
   Call gs_LimpiaGrid(grd_LisOpe)
End Sub

Private Sub cmd_BusOpe_Click()
   Dim r_int_FlgEnc  As Integer

   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   r_int_FlgEnc = 0
   
   grd_LisOpe.Redraw = False
   
   Call gs_LimpiaGrid(grd_LisOpe)
   
   'Buscando Cliente
   g_str_Parame = "SELECT * FROM CTB_DETITF WHERE "
   g_str_Parame = g_str_Parame & "DETITF_TIPDOC = " & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "DETITF_NUMDOC = '" & Trim(txt_NumDoc.Text) & "' "
   g_str_Parame = g_str_Parame & "ORDER BY DETITF_PERANO DESC, DETITF_PERMES DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgEnc = 1
      
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_LisOpe.Rows = grd_LisOpe.Rows + 1
         grd_LisOpe.Row = grd_LisOpe.Rows - 1
         
         grd_LisOpe.Col = 0
         grd_LisOpe.Text = Trim(g_rst_Princi!DETITF_PERANO) & "-" & Format(Trim(g_rst_Princi!DETITF_PERMES), "00")
         
         grd_LisOpe.Col = 1
         grd_LisOpe.Text = IIf(Trim(g_rst_Princi!DETITF_TIPDEC) = 1, "DECLARANTE", "EXTORNO")
         
         grd_LisOpe.Col = 2
         grd_LisOpe.Text = gf_FormatoFecha(g_rst_Princi!DETITF_FECMOV)
         
         grd_LisOpe.Col = 3
         grd_LisOpe.Text = Trim(g_rst_Princi!DETITF_TIPMOV)
         
         grd_LisOpe.Col = 4
         grd_LisOpe.Text = Format(g_rst_Princi!DETITF_MTOSOL, "###,###,###,##0.00")
         
         grd_LisOpe.Col = 5
         grd_LisOpe.Text = Format(g_rst_Princi!DETITF_ITFSOL, "###,###,###,##0.00")
         
         grd_LisOpe.Col = 6
         grd_LisOpe.Text = Trim(g_rst_Princi!DETITF_MANUAL)
         
         grd_LisOpe.Col = 7
         grd_LisOpe.Text = Trim(g_rst_Princi!DETITF_TIPDEC)
                  
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
     
   grd_LisOpe.Redraw = True
   
   If grd_LisOpe.Rows > 0 Then
      'Call pnl_Tit_NumOpe_Click
      
      Call gs_UbiIniGrid(grd_LisOpe)
   Else
      MsgBox "No se encontró ningún registro para este Documento de Identidad.", vbInformation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub grd_LisOpe_SelChange()
   If grd_LisOpe.Rows > 2 Then
      grd_LisOpe.RowSel = grd_LisOpe.Row
   End If
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_BusOpe)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 5:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 6:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
         End Select
      Else
         KeyAscii = 0
      End If
   End If

End Sub

