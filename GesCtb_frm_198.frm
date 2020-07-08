VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_InvDpf_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15420
   Icon            =   "GesCtb_frm_198.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   7365
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15825
      _Version        =   65536
      _ExtentX        =   27914
      _ExtentY        =   12991
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
         TabIndex        =   1
         Top             =   60
         Width           =   15330
         _Version        =   65536
         _ExtentX        =   27040
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
            Caption         =   "Detalle de la Inversión del Depósito Plazo Fijo"
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
            Picture         =   "GesCtb_frm_198.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   3
         Top             =   780
         Width           =   15330
         _Version        =   65536
         _ExtentX        =   27040
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_198.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Consul 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_198.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consultar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14700
            Picture         =   "GesCtb_frm_198.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   5475
         Left            =   60
         TabIndex        =   7
         Top             =   1470
         Width           =   15330
         _Version        =   65536
         _ExtentX        =   27040
         _ExtentY        =   9657
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
            Height          =   5085
            Left            =   30
            TabIndex        =   8
            Top             =   360
            Width           =   15290
            _ExtentX        =   26961
            _ExtentY        =   8969
            _Version        =   393216
            Rows            =   24
            Cols            =   13
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_DebMN 
            Height          =   285
            Left            =   5370
            TabIndex        =   9
            Top             =   60
            Width           =   900
            _Version        =   65536
            _ExtentX        =   1587
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Plazo Dias"
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
         Begin Threed.SSPanel pnl_HabME 
            Height          =   285
            Left            =   6240
            TabIndex        =   10
            Top             =   60
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1402
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tasa %"
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   7020
            TabIndex        =   11
            Top             =   60
            Width           =   800
            _Version        =   65536
            _ExtentX        =   1411
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   7800
            TabIndex        =   12
            Top             =   60
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Capital"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   1080
            TabIndex        =   13
            Top             =   60
            Width           =   2500
            _Version        =   65536
            _ExtentX        =   4410
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Institución"
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
            Left            =   3570
            TabIndex        =   14
            Top             =   60
            Width           =   1810
            _Version        =   65536
            _ExtentX        =   3193
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Operación de Ref."
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1065
            _Version        =   65536
            _ExtentX        =   1887
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro Cuenta"
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
            Left            =   10440
            TabIndex        =   16
            Top             =   60
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1976
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vencimiento"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   11550
            TabIndex        =   17
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2152
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Rendimeinto"
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   285
            Left            =   12750
            TabIndex        =   18
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2152
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Devengado"
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   9315
            TabIndex        =   19
            Top             =   60
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Apertura"
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
         Begin Threed.SSPanel SSPanel17 
            Height          =   285
            Left            =   13950
            TabIndex        =   20
            Top             =   60
            Width           =   1020
            _Version        =   65536
            _ExtentX        =   1799
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Estado"
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
   End
End
Attribute VB_Name = "frm_Ctb_InvDpf_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim r_aux_Codigo     As String
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1020 'NRO CUENTA
   grd_Listad.ColWidth(1) = 2490 'INSITUCION
   grd_Listad.ColWidth(2) = 1800 'OPERACION DE REF.
   grd_Listad.ColWidth(3) = 870 'PLAZO
   grd_Listad.ColWidth(4) = 780 'TASA
   grd_Listad.ColWidth(5) = 800 'MONEDA
   grd_Listad.ColWidth(6) = 1500 'CAPITAL
   grd_Listad.ColWidth(7) = 1120 'FECHA APERTURA
   grd_Listad.ColWidth(8) = 1110 'FECHA VECIMIENTO
   grd_Listad.ColWidth(9) = 1190 'RENDIMIENTO
   grd_Listad.ColWidth(10) = 1210 'DEVENGADO
   grd_Listad.ColWidth(11) = 1010 'NOM_SITUACION
   grd_Listad.ColWidth(12) = 0 'COD_SITUACION
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignCenterCenter
   grd_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub cmd_Consul_Click()
   moddat_g_str_Codigo = ""
   moddat_g_int_Situac = 0
   moddat_g_str_Situac = ""
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
 
   Call gs_RefrescaGrid(grd_Listad)
   
   grd_Listad.Col = 0
   moddat_g_str_Codigo = CStr(grd_Listad.Text)
   grd_Listad.Col = 11
   moddat_g_str_Situac = CStr(grd_Listad.Text) 'Descripcion
   grd_Listad.Col = 12
   moddat_g_int_Situac = CInt(grd_Listad.Text) 'Codigo
   
   Call gs_RefrescaGrid(grd_Listad)
   'moddat_g_int_InsAct = 0
   moddat_g_int_FlgGrb = 0
   frm_Ctb_InvDpf_02.Show 1
   
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub fs_Buscar()
Dim r_aux_Codigo     As String
Dim r_str_FecVct     As String
Dim r_str_FecApe     As String
Dim r_int_FecDif     As Integer

   r_aux_Codigo = CStr(CLng(moddat_g_str_Codigo)) & ","
   
   Do While CLng(moddat_g_str_Codigo) > 0
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT A.MAEDPF_NUMCTA_REF  "
      g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEDPF A  "
      g_str_Parame = g_str_Parame & "  WHERE A.MAEDPF_NUMCTA =  " & CLng(moddat_g_str_Codigo)
      g_str_Parame = g_str_Parame & "    AND MAEDPF_SITUAC = 1 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         'MsgBox "No se ha encontrado ningún registro.", vbExclamation, modgen_g_str_NomPlt
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      Else
         g_rst_Princi.MoveFirst
         If Len(Trim(g_rst_Princi!MAEDPF_NUMCTA_REF & "")) = 0 Then
             moddat_g_str_Codigo = "0"
             Exit Do
         Else
             r_aux_Codigo = r_aux_Codigo & g_rst_Princi!MAEDPF_NUMCTA_REF & ","
             moddat_g_str_Codigo = Trim(g_rst_Princi!MAEDPF_NUMCTA_REF & "")
         End If
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         Screen.MousePointer = 0
      End If
   Loop
   r_aux_Codigo = "(" & Mid(r_aux_Codigo, 1, Len(r_aux_Codigo) - 1) & ")"
   '-------------------------------------1111111--------------------------------
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT LPAD(A.MAEDPF_NUMCTA,8,'0') MAEDPF_NUMCTA, TRIM(B.PARDES_DESCRI) AS ENTIDAD_DEST, A.MAEDPF_NUMREF, A.MAEDPF_PLADIA,  "
   g_str_Parame = g_str_Parame & "        A.MAEDPF_TASINT, A.MAEDPF_CODMON, A.MAEDPF_SALCAP, A.MAEDPF_INTAJU,  "
   g_str_Parame = g_str_Parame & "        A.MAEDPF_INTCAP, TRIM(C.PARDES_DESCRI) AS ENTIDAD_ORIG, TRIM(D.PARDES_DESCRI) AS MONEDA,  "
   g_str_Parame = g_str_Parame & "        A.MAEDPF_FECAPE, TO_CHAR(TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA,'yyyymmdd') AS FEC_VCTO,  "
   g_str_Parame = g_str_Parame & "        DECODE(A.MAEDPF_TIPDPF,1,  "
   g_str_Parame = g_str_Parame & "               CASE  "
   g_str_Parame = g_str_Parame & "               WHEN (TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA) <= TO_DATE(SYSDATE,'DD/MM/YY') THEN 'VENCIDO'  "
   g_str_Parame = g_str_Parame & "                    WHEN TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA > TO_DATE(SYSDATE,'DD/MM/YY') THEN 'VIGENTE'  "
   g_str_Parame = g_str_Parame & "                  END, 'CERRADO') AS NOM_SITUAC,  "
   g_str_Parame = g_str_Parame & "           DECODE(A.MAEDPF_TIPDPF,1,  "
   g_str_Parame = g_str_Parame & "           CASE  "
   g_str_Parame = g_str_Parame & "             WHEN (TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA) <= TO_DATE(SYSDATE,'DD/MM/YY') THEN 2  "
   g_str_Parame = g_str_Parame & "             WHEN TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA > TO_DATE(SYSDATE,'DD/MM/YY') THEN 1  "
   g_str_Parame = g_str_Parame & "           END, A.MAEDPF_SITDPF) AS COD_SITUAC  "
   g_str_Parame = g_str_Parame & "      FROM CNTBL_MAEDPF A  "
   g_str_Parame = g_str_Parame & "     INNER JOIN MNT_PARDES B ON A.MAEDPF_CODENT_DES = B.PARDES_CODITE AND B.PARDES_CODGRP = 122  "
   g_str_Parame = g_str_Parame & "     INNER JOIN MNT_PARDES C ON A.MAEDPF_CODENT_ORI = C.PARDES_CODITE AND C.PARDES_CODGRP = 122  "
   g_str_Parame = g_str_Parame & "     INNER JOIN MNT_PARDES D ON A.MAEDPF_CODMON = D.PARDES_CODITE AND D.PARDES_CODGRP = 204    "
   g_str_Parame = g_str_Parame & "     WHERE A.MAEDPF_SITUAC = 1  "
   g_str_Parame = g_str_Parame & "       AND A.MAEDPF_NUMCTA IN  " & r_aux_Codigo
   g_str_Parame = g_str_Parame & "     ORDER BY A.MAEDPF_NUMCTA ASC  "
   
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
   
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = CStr(g_rst_Princi!MAEDPF_NUMCTA)
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!ENTIDAD_DEST & "")
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!MAEDPF_NUMREF & "")
      
      grd_Listad.Col = 3
      grd_Listad.Text = g_rst_Princi!MAEDPF_PLADIA
      
      grd_Listad.Col = 4
      grd_Listad.Text = Format(Trim(g_rst_Princi!MAEDPF_TASINT & ""), "###,###,###,##0.00")
      
      grd_Listad.Col = 5
      grd_Listad.Text = Trim(g_rst_Princi!Moneda & "")
                                    
      grd_Listad.Col = 6
      grd_Listad.Text = Format(g_rst_Princi!MAEDPF_SALCAP, "###,###,###,##0.00")

      grd_Listad.Col = 7
      grd_Listad.Text = gf_FormatoFecha(Trim(g_rst_Princi!MAEDPF_FECAPE & ""))
      r_str_FecApe = grd_Listad.Text
      
      grd_Listad.Col = 8
      grd_Listad.Text = DateAdd("d", g_rst_Princi!MAEDPF_PLADIA, gf_FormatoFecha(Trim(g_rst_Princi!MAEDPF_FECAPE & "")))
      r_str_FecVct = grd_Listad.Text

      grd_Listad.Col = 9
      'grd_Listad.Text = Format(CDbl(g_rst_Princi!MAEDPF_INTCAP) + CDbl(g_rst_Princi!MAEDPF_INTAJU), "###,###,###,##0.00")
      grd_Listad.Text = Format(CDbl(g_rst_Princi!MAEDPF_INTCAP), "###,###,###,##0.00")
                        
      '=SI(F_DIA=F_VCTO,INTCAP, SI(F_DIA>F_VCTO,INTCAP, ((((1+TASA)^(1/360))-1)*CAPITAL)*(F_DIA-F_APER+1) ))
      grd_Listad.Col = 10
      grd_Listad.Text = "0.00"
      r_str_FecApe = Format(r_str_FecApe, "yyyymmdd")
      r_str_FecVct = Format(r_str_FecVct, "yyyymmdd")
      r_int_FecDif = 0
'      If Format(moddat_g_str_FecSis, "yyyymmdd") = r_str_FecVct Then
'         grd_Listad.Text = Format(CDbl(g_rst_Princi!MAEDPF_INTCAP) + CDbl(g_rst_Princi!MAEDPF_INTAJU), "###,###,###,##0.00")
'      ElseIf Format(moddat_g_str_FecSis, "yyyymmdd") > r_str_FecVct Then
'         grd_Listad.Text = Format(CDbl(g_rst_Princi!MAEDPF_INTCAP) + CDbl(g_rst_Princi!MAEDPF_INTAJU), "###,###,###,##0.00")
'      Else
'         r_int_FecDif = DateDiff("d", gf_FormatoFecha(r_str_FecApe), moddat_g_str_FecSis) + 1
'         grd_Listad.Text = ((((1 + ((CDbl(g_rst_Princi!MAEDPF_TASINT) + CDbl(g_rst_Princi!MAEDPF_INTAJU)) / 100)) ^ (1 / 360)) - 1) * g_rst_Princi!MAEDPF_SALCAP) * r_int_FecDif
'         grd_Listad.Text = Format(grd_Listad.Text, "###,###,###,##0.00")
'      End If
      Dim ImpAux As String
      Call frm_Ctb_InvDpf_01.fs_CalDevengado(moddat_g_str_FecSis, r_str_FecVct, CDbl(g_rst_Princi!MAEDPF_INTCAP), _
                           r_str_FecApe, g_rst_Princi!MAEDPF_TASINT, g_rst_Princi!MAEDPF_SALCAP, ImpAux)
      grd_Listad.Text = ImpAux
            
      grd_Listad.Col = 11
      grd_Listad.Text = Trim(g_rst_Princi!NOM_SITUAC & "")
      
      grd_Listad.Col = 12
      grd_Listad.Text = Trim(g_rst_Princi!COD_SITUAC & "")
      
      g_rst_Princi.MoveNext
   Loop

   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_int_NumFil        As Integer
Dim r_int_Contar        As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE INVERSIONES - DEPOSITO PLAZO FIJO"
      .Range(.Cells(2, 2), .Cells(2, 13)).Merge
      .Range(.Cells(2, 2), .Cells(2, 13)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 13)).HorizontalAlignment = xlHAlignCenter

      .Cells(3, 2) = "NRO CUENTA"
      .Cells(3, 3) = "INSTITUCION"
      .Cells(3, 4) = "OPERACIONE DE REF."
      .Cells(3, 5) = "PLAZO DIAS"
      .Cells(3, 6) = "TASA %"
      .Cells(3, 7) = "MONEDA"
      .Cells(3, 8) = "CAPITAL"
      .Cells(3, 9) = "F.APERTURA"
      .Cells(3, 10) = "F.VENCIMIENTO"
      .Cells(3, 11) = "RENDIMIENTO"
      .Cells(3, 12) = "DEVENGADO"
      .Cells(3, 13) = "ESTADO"
         
      .Range(.Cells(3, 2), .Cells(3, 13)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(3, 2), .Cells(3, 13)).Font.Bold = True
       
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 12 'Nro Cuenta
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 30 'institucion
      .Columns("C").HorizontalAlignment = xlHAlignLeft
      .Columns("D").ColumnWidth = 23 'Operacion de Ref.
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").ColumnWidth = 11 'Plazo Dias
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 10 'Tasa %
      '.Columns("F").NumberFormat = "0.00%"
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 10 'Moneda
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 16 'Capital
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 13 'F.Apertura
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 15 'F.Vencimiento
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 16 'Rendimiento
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 16 'Devengado
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 12 'Estado
      .Columns("M").HorizontalAlignment = xlHAlignCenter
            
      .Range(.Cells(1, 1), .Cells(10, 13)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(10, 13)).Font.Size = 11
      
      r_int_NumFil = 2
      For r_int_Contar = 0 To grd_Listad.Rows - 1
          .Cells(r_int_NumFil + 2, 2) = "'" & grd_Listad.TextMatrix(r_int_Contar, 0) 'Nro Cuenta
          .Cells(r_int_NumFil + 2, 3) = "'" & grd_Listad.TextMatrix(r_int_Contar, 1) 'institucion
          .Cells(r_int_NumFil + 2, 4) = "'" & grd_Listad.TextMatrix(r_int_Contar, 2) 'Operacion de Ref.
          .Cells(r_int_NumFil + 2, 5) = "'" & grd_Listad.TextMatrix(r_int_Contar, 3) 'Plazo Dias
          .Cells(r_int_NumFil + 2, 6) = "'" & grd_Listad.TextMatrix(r_int_Contar, 4) 'Tasa %
          .Cells(r_int_NumFil + 2, 7) = "'" & grd_Listad.TextMatrix(r_int_Contar, 5) 'Moneda
          .Cells(r_int_NumFil + 2, 8) = grd_Listad.TextMatrix(r_int_Contar, 6)  'Capital
          .Cells(r_int_NumFil + 2, 9) = "'" & grd_Listad.TextMatrix(r_int_Contar, 7) 'F.Apertura
          .Cells(r_int_NumFil + 2, 10) = "'" & grd_Listad.TextMatrix(r_int_Contar, 8) 'F.Vencimiento
          .Cells(r_int_NumFil + 2, 11) = grd_Listad.TextMatrix(r_int_Contar, 9) 'Rendimiento
          .Cells(r_int_NumFil + 2, 12) = grd_Listad.TextMatrix(r_int_Contar, 10)  'Devengado
          .Cells(r_int_NumFil + 2, 13) = "'" & grd_Listad.TextMatrix(r_int_Contar, 11) 'Estado
          
          r_int_NumFil = r_int_NumFil + 1
      Next
      
      .Range(.Cells(3, 3), .Cells(3, 13)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
