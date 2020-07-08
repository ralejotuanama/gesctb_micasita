VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Mnt_SecEco_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   2685
   ClientTop       =   2550
   ClientWidth     =   13260
   Icon            =   "GesCtb_frm_136.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8055
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13245
      _Version        =   65536
      _ExtentX        =   23363
      _ExtentY        =   14208
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   60
         Width           =   13125
         _Version        =   65536
         _ExtentX        =   23151
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   480
            Left            =   630
            TabIndex        =   5
            Top             =   60
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   847
            _StockProps     =   15
            Caption         =   "Mantenimiento de Sectores Económicos"
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
            Picture         =   "GesCtb_frm_136.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   6
         Top             =   780
         Width           =   13125
         _Version        =   65536
         _ExtentX        =   23151
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
            Left            =   60
            Picture         =   "GesCtb_frm_136.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12510
            Picture         =   "GesCtb_frm_136.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   435
         Left            =   30
         TabIndex        =   7
         Top             =   1470
         Width           =   13125
         _Version        =   65536
         _ExtentX        =   23151
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
         Begin Threed.SSPanel pnl_SecEco 
            Height          =   315
            Left            =   1650
            TabIndex        =   11
            Top             =   60
            Width           =   11415
            _Version        =   65536
            _ExtentX        =   20135
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Label Label4 
            Caption         =   "Sector Económico:"
            Height          =   255
            Left            =   60
            TabIndex        =   8
            Top             =   60
            Width           =   1665
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   6045
         Left            =   30
         TabIndex        =   9
         Top             =   1950
         Width           =   13125
         _Version        =   65536
         _ExtentX        =   23151
         _ExtentY        =   10663
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
         Begin MSComctlLib.TreeView tvw_LisCiu 
            Height          =   5625
            Left            =   60
            TabIndex        =   0
            Top             =   360
            Width           =   13005
            _ExtentX        =   22939
            _ExtentY        =   9922
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   2
            Checkboxes      =   -1  'True
            SingleSel       =   -1  'True
            Appearance      =   1
         End
         Begin Threed.SSPanel pnl_Tit_Descri 
            Height          =   285
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   13005
            _Version        =   65536
            _ExtentX        =   22939
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Relación de CIIU"
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
Attribute VB_Name = "frm_Mnt_SecEco_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Grabar_Click()
   Dim r_str_CodCiu     As String
   Dim r_int_Contad     As Integer
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
      Exit Sub
   End If
   
   'Borrando CIIU para Sector Económico
   g_str_Parame = "DELETE FROM MNT_SECCIU WHERE "
   g_str_Parame = g_str_Parame & "SECCIU_CODSEC = '" & moddat_g_str_CodGrp & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   'Agregando CIIU
   For r_int_Contad = 1 To tvw_LisCiu.Nodes.Count
      If tvw_LisCiu.Nodes(r_int_Contad).Checked = True Then
         r_str_CodCiu = Mid(tvw_LisCiu.Nodes(r_int_Contad).Text, 1, 4)
         
         moddat_g_int_FlgGOK = False
         moddat_g_int_CntErr = 0
         
         Do While moddat_g_int_FlgGOK = False
            g_str_Parame = "USP_MNT_SECCIU ("
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodGrp & "', "
            g_str_Parame = g_str_Parame & "'" & r_str_CodCiu & "', "
               
            'Datos de Auditoria
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
               
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               moddat_g_int_CntErr = moddat_g_int_CntErr + 1
            Else
               moddat_g_int_FlgGOK = True
            End If
      
            If moddat_g_int_CntErr = 6 Then
               If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Exit Sub
               Else
                  moddat_g_int_CntErr = 0
               End If
            End If
         Loop
      End If
   Next r_int_Contad
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt
   
   pnl_SecEco.Caption = moddat_g_str_DesGrp

   Call gs_CentraForm(Me)
   
   Call fs_Buscar

   Screen.MousePointer = 0
End Sub

Private Sub fs_Buscar()
   Dim r_str_CodNv1     As String
   Dim r_str_CodNv2     As String
   Dim r_obj_NodNvo     As Node
   Dim r_rst_Genera     As ADODB.Recordset
   Dim r_str_CodCiu     As String
      
   tvw_LisCiu.Nodes.Clear
   
   g_str_Parame = "SELECT * FROM MNT_PARDES WHERE PARDES_CODGRP = '102' AND PARDES_CODITE <> '000000' ORDER BY PARDES_CODITE ASC"
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
     Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      r_str_CodCiu = Mid(g_rst_Princi!PARDES_CODITE, 3, 4)
      
      g_str_Parame = "SELECT * FROM MNT_SECCIU WHERE SECCIU_CODCIU = " & Mid(g_rst_Princi!PARDES_CODITE, 3, 4)
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Genera.EOF And g_rst_Genera.BOF) Then
         g_str_Parame = "SELECT * FROM MNT_SECCIU WHERE SECCIU_CODSEC = " & moddat_g_str_CodGrp & " AND SECCIU_CODCIU = " & r_str_CodCiu
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
            Set r_obj_NodNvo = tvw_LisCiu.Nodes.Add(, , "A" + r_str_CodCiu, r_str_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", g_rst_Princi!PARDES_CODITE))
            tvw_LisCiu.Nodes(tvw_LisCiu.Nodes.Count).Checked = True
         End If
         
         r_rst_Genera.Close
         Set r_rst_Genera = Nothing
      Else
         Set r_obj_NodNvo = tvw_LisCiu.Nodes.Add(, , "A" + r_str_CodCiu, r_str_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", g_rst_Princi!PARDES_CODITE))
         tvw_LisCiu.Nodes(tvw_LisCiu.Nodes.Count).Checked = False
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If tvw_LisCiu.Nodes.Count = 0 Then
      cmd_Grabar.Enabled = False
   Else
      cmd_Grabar.Enabled = True
   End If
End Sub

