VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_RptCtb_29 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14100
   Icon            =   "GesCtb_frm_911.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   14100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   14085
      _Version        =   65536
      _ExtentX        =   24844
      _ExtentY        =   12303
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
         TabIndex        =   9
         Top             =   60
         Width           =   13965
         _Version        =   65536
         _ExtentX        =   24633
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   270
            Left            =   630
            TabIndex        =   10
            Top             =   150
            Width           =   5205
            _Version        =   65536
            _ExtentX        =   9181
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Consolidado de Cartera en Riesgo"
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
            Picture         =   "GesCtb_frm_911.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   11
         Top             =   780
         Width           =   13965
         _Version        =   65536
         _ExtentX        =   24633
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
            Left            =   13350
            Picture         =   "GesCtb_frm_911.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_911.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_911.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Procesar informacion"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   885
         Left            =   60
         TabIndex        =   12
         Top             =   1470
         Width           =   13965
         _Version        =   65536
         _ExtentX        =   24633
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
         Begin VB.ComboBox cmb_TipDat 
            Height          =   315
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   480
            Width           =   2265
         End
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   5460
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   135
            Width           =   2265
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   135
            Width           =   2265
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1170
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
         Begin VB.Label Label4 
            Caption         =   "Expresado en:"
            Height          =   315
            Left            =   4050
            TabIndex        =   17
            Top             =   510
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Reporte:"
            Height          =   315
            Left            =   4050
            TabIndex        =   16
            Top             =   195
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   180
            TabIndex        =   14
            Top             =   510
            Width           =   885
         End
         Begin VB.Label Label2 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   180
            TabIndex        =   13
            Top             =   195
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   4560
         Left            =   60
         TabIndex        =   15
         Top             =   2400
         Width           =   13965
         _Version        =   65536
         _ExtentX        =   24633
         _ExtentY        =   8043
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisCla 
            Height          =   4455
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   13830
            _ExtentX        =   24395
            _ExtentY        =   7858
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            Redraw          =   -1  'True
            MergeCells      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
Attribute VB_Name = "frm_RptCtb_29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExc_Click()
   If cmb_PerMes.ListIndex = -1 Then
       MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExcRes
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Proces_Click()
Dim r_int_MesAct     As Integer
Dim r_int_AnoAct     As Integer
Dim r_int_TipRep     As Integer
Dim r_int_TipExp     As Integer
   
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el mes a consultar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe ingresar el año a consultar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de reporte a consultar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If cmb_TipDat.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de dato en que se expresara el reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDat)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   r_int_MesAct = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   r_int_AnoAct = ipp_PerAno.Text
   r_int_TipRep = cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
   r_int_TipExp = cmb_TipDat.ItemData(cmb_TipDat.ListIndex)
   
   If r_int_TipRep = 1 Then
      Call fs_Setea_Columnas
      Call fs_Reporte_Interno(r_int_MesAct, r_int_AnoAct, r_int_TipExp)
   Else
      Call fs_Setea_Columnas
      Call fs_Reporte_Opic(r_int_MesAct, r_int_AnoAct, r_int_TipExp)
   End If
   
   cmd_ExpExc.Enabled = True
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
    
   Call gs_CentraForm(Me)
   Call fs_Inicia
     
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
   cmd_ExpExc.Enabled = False
    
   cmb_TipRep.Clear
   cmb_TipRep.AddItem "INTERNO - SBS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   cmb_TipRep.AddItem "OPIC"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   cmb_TipRep.ListIndex = -1

   cmb_TipDat.Clear
   cmb_TipDat.AddItem "MONTOS"
   cmb_TipDat.ItemData(cmb_TipDat.NewIndex) = 1
   cmb_TipDat.AddItem "NUMEROS"
   cmb_TipDat.ItemData(cmb_TipDat.NewIndex) = 2
   cmb_TipDat.ListIndex = -1
End Sub

Private Sub fs_Setea_Columnas()
   grd_LisCla.Redraw = False
   Call gs_LimpiaGrid(grd_LisCla)
   
   'Ancho de columnas
   grd_LisCla.Cols = 14
   grd_LisCla.ColWidth(0) = 0
   grd_LisCla.ColWidth(1) = 4200
   grd_LisCla.ColWidth(2) = 1300
   grd_LisCla.ColWidth(3) = 1300
   grd_LisCla.ColWidth(4) = 1300
   grd_LisCla.ColWidth(5) = 1300
   grd_LisCla.ColWidth(6) = 1300
   grd_LisCla.ColWidth(7) = 1300
   grd_LisCla.ColWidth(8) = 1300
   grd_LisCla.ColWidth(9) = 1300
   grd_LisCla.ColWidth(10) = 1300
   grd_LisCla.ColWidth(11) = 1300
   grd_LisCla.ColWidth(12) = 1300
   grd_LisCla.ColWidth(13) = 1300
   grd_LisCla.ColAlignment(1) = flexAlignLeftCenter
   grd_LisCla.ColAlignment(2) = flexAlignRightCenter
   grd_LisCla.ColAlignment(3) = flexAlignRightCenter
   grd_LisCla.ColAlignment(4) = flexAlignRightCenter
   grd_LisCla.ColAlignment(5) = flexAlignRightCenter
   grd_LisCla.ColAlignment(6) = flexAlignRightCenter
   grd_LisCla.ColAlignment(7) = flexAlignRightCenter
   grd_LisCla.ColAlignment(8) = flexAlignRightCenter
   grd_LisCla.ColAlignment(9) = flexAlignRightCenter
   grd_LisCla.ColAlignment(10) = flexAlignRightCenter
   grd_LisCla.ColAlignment(11) = flexAlignRightCenter
   grd_LisCla.ColAlignment(12) = flexAlignRightCenter
   grd_LisCla.ColAlignment(13) = flexAlignRightCenter

   'Cabecera
   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Row = 0: grd_LisCla.Text = ""
   grd_LisCla.Col = 1: grd_LisCla.Text = ""
   grd_LisCla.Col = 2: grd_LisCla.Text = "ENERO":        grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 3: grd_LisCla.Text = "FEBRERO":      grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 4: grd_LisCla.Text = "MARZO":        grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 5: grd_LisCla.Text = "ABRIL":        grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 6: grd_LisCla.Text = "MAYO":         grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 7: grd_LisCla.Text = "JUNIO":        grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 8: grd_LisCla.Text = "JULIO":        grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 9: grd_LisCla.Text = "AGOSTO":       grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 10: grd_LisCla.Text = "SETIEMBRE":   grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 11: grd_LisCla.Text = "OCTUBRE":     grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 12: grd_LisCla.Text = "NOVIEMBRE":   grd_LisCla.CellAlignment = flexAlignCenterCenter
   grd_LisCla.Col = 13: grd_LisCla.Text = "DICIEMBRE":   grd_LisCla.CellAlignment = flexAlignCenterCenter

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "0"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera Vigente"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "1"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera en riesgo (1-30 dias)"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "2"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera en riesgo (> 30 dias)"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "3"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera en riesgo (30-60 dias)"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "4"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera en riesgo (60-90 dias)"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "5"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera en riesgo (90-120 dias)"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "6"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera en riesgo (120-180 dias)"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "7"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera en riesgo (180-360 dias)"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "8"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera en riesgo (> 360 dias)"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "9"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "*no incluyendo cartera reestructurada y/o reprogramada"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "10"
   grd_LisCla.Col = 1:   grd_LisCla.Text = ""

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "11"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera reestructurada y/o reprogramada"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "12"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera Vigente"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "13"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera en riesgo (1-30 dias)"

   grd_LisCla.Rows = grd_LisCla.Rows + 1
   grd_LisCla.Row = grd_LisCla.Rows - 1
   grd_LisCla.Col = 0:   grd_LisCla.Text = "13"
   grd_LisCla.Col = 1:   grd_LisCla.Text = "Cartera en riesgo (> 30 dias)"

   With grd_LisCla
      .MergeCells = flexMergeFree
      .FixedCols = 2
      .FixedRows = 1
   End With
   grd_LisCla.Redraw = True
End Sub

Private Sub fs_Reporte_Interno(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_TipExp As Integer)
Dim r_int_Col        As Integer
Dim r_int_Count      As Integer
Dim r_int_MesSel     As Integer

   r_int_Col = 2
   For r_int_Count = 1 To p_PerMes
      If r_int_Count > p_PerMes Then
         Exit For
      End If
      
      r_int_MesSel = r_int_Count
      
      If p_TipExp = 1 Then
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " SELECT "
         g_str_Parame = g_str_Parame & "     ((SELECT NVL(ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, HIPCIE_CAPVIG, (HIPCIE_CAPVIG*HIPCIE_TIPCAM))), 2), 0) "
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0) "
         g_str_Parame = g_str_Parame & "        + "
         g_str_Parame = g_str_Parame & "      (SELECT NVL(ROUND(SUM(DECODE(COMCIE_TIPMON, 1, COMCIE_CAPVIG, (COMCIE_CAPVIG*COMCIE_TIPCAM))), 2), 0) "
         g_str_Parame = g_str_Parame & "         FROM CRE_COMCIE"
         g_str_Parame = g_str_Parame & "        WHERE COMCIE_PERMES = " & r_int_MesSel & " AND COMCIE_PERANO = " & p_PerAno & ")) AS CAR_VIG_01,"
         g_str_Parame = g_str_Parame & "      (SELECT NVL(ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN), (HIPCIE_CAPVEN*HIPCIE_TIPCAM))), 0), 2), 0) "
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 30 AND HIPCIE_DIAMOR <= 60)   AS CAR_VIG_02,"
         g_str_Parame = g_str_Parame & "      (SELECT NVL(ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN), (HIPCIE_CAPVEN*HIPCIE_TIPCAM))), 0), 2), 0) "
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 60 AND HIPCIE_DIAMOR <= 90)   AS CAR_VIG_03,"
         g_str_Parame = g_str_Parame & "      (SELECT NVL(ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN), (HIPCIE_CAPVEN*HIPCIE_TIPCAM))), 0), 2), 0) "
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 90 AND HIPCIE_DIAMOR <= 120)  AS CAR_VIG_04,"
         g_str_Parame = g_str_Parame & "      (SELECT NVL(ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN), (HIPCIE_CAPVEN*HIPCIE_TIPCAM))), 0), 2), 0)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 120 AND HIPCIE_DIAMOR <= 180) AS CAR_VIG_05,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN), (HIPCIE_CAPVEN*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 180 AND HIPCIE_DIAMOR <= 360) AS CAR_VIG_06,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN), (HIPCIE_CAPVEN*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 360)                          AS CAR_VIG_07,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVIG), (HIPCIE_CAPVIG*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 1)   AS CAR_REP_01,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN), (HIPCIE_CAPVEN*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 1"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 30)                                                                                    AS CAR_REP_02"
         g_str_Parame = g_str_Parame & "  FROM DUAL"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Screen.MousePointer = 0
            Exit Sub
         End If
         
         grd_LisCla.TextMatrix(1, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_01), 0, g_rst_Princi!CAR_VIG_01), "###,###,##0.00")
         grd_LisCla.TextMatrix(2, r_int_Col) = ""
         grd_LisCla.TextMatrix(3, r_int_Col) = ""
         grd_LisCla.TextMatrix(4, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_02), 0, g_rst_Princi!CAR_VIG_02), "###,###,##0.00")
         grd_LisCla.TextMatrix(5, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_03), 0, g_rst_Princi!CAR_VIG_03), "###,###,##0.00")
         grd_LisCla.TextMatrix(6, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_04), 0, g_rst_Princi!CAR_VIG_04), "###,###,##0.00")
         grd_LisCla.TextMatrix(7, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_05), 0, g_rst_Princi!CAR_VIG_05), "###,###,##0.00")
         grd_LisCla.TextMatrix(8, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_06), 0, g_rst_Princi!CAR_VIG_06), "###,###,##0.00")
         grd_LisCla.TextMatrix(9, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_07), 0, g_rst_Princi!CAR_VIG_07), "###,###,##0.00")
         grd_LisCla.TextMatrix(10, r_int_Col) = ""
         grd_LisCla.TextMatrix(11, r_int_Col) = ""
         grd_LisCla.TextMatrix(12, r_int_Col) = ""
         grd_LisCla.TextMatrix(13, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_REP_01), 0, g_rst_Princi!CAR_REP_01), "###,###,##0.00")
         grd_LisCla.TextMatrix(14, r_int_Col) = ""
         grd_LisCla.TextMatrix(15, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_REP_02), 0, g_rst_Princi!CAR_REP_02), "###,###,##0.00")
      
      Else
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " SELECT "
         g_str_Parame = g_str_Parame & "     ((SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0) "
         g_str_Parame = g_str_Parame & "        + "
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_COMCIE"
         g_str_Parame = g_str_Parame & "        WHERE COMCIE_PERMES = " & r_int_MesSel & " AND COMCIE_PERANO = " & p_PerAno & ")) AS CAR_VIG_01,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 30 AND HIPCIE_DIAMOR <= 60)   AS CAR_VIG_02,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 60 AND HIPCIE_DIAMOR <= 90)   AS CAR_VIG_03,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 90 AND HIPCIE_DIAMOR <= 120)  AS CAR_VIG_04,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 120 AND HIPCIE_DIAMOR <= 180) AS CAR_VIG_05,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 180 AND HIPCIE_DIAMOR <= 360) AS CAR_VIG_06,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 360)                          AS CAR_VIG_07,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 1)   AS CAR_REP_01,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 1"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 30)                                                                                    AS CAR_REP_02"
         g_str_Parame = g_str_Parame & "  FROM DUAL"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Screen.MousePointer = 0
            Exit Sub
         End If
         
         grd_LisCla.TextMatrix(1, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_01), 0, g_rst_Princi!CAR_VIG_01), "##,##0")
         grd_LisCla.TextMatrix(2, r_int_Col) = ""
         grd_LisCla.TextMatrix(3, r_int_Col) = ""
         grd_LisCla.TextMatrix(4, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_02), 0, g_rst_Princi!CAR_VIG_02), "##,##0")
         grd_LisCla.TextMatrix(5, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_03), 0, g_rst_Princi!CAR_VIG_03), "##,##0")
         grd_LisCla.TextMatrix(6, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_04), 0, g_rst_Princi!CAR_VIG_04), "##,##0")
         grd_LisCla.TextMatrix(7, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_05), 0, g_rst_Princi!CAR_VIG_05), "##,##0")
         grd_LisCla.TextMatrix(8, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_06), 0, g_rst_Princi!CAR_VIG_06), "##,##0")
         grd_LisCla.TextMatrix(9, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_07), 0, g_rst_Princi!CAR_VIG_07), "##,##0")
         grd_LisCla.TextMatrix(10, r_int_Col) = ""
         grd_LisCla.TextMatrix(11, r_int_Col) = ""
         grd_LisCla.TextMatrix(12, r_int_Col) = ""
         grd_LisCla.TextMatrix(13, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_REP_01), 0, g_rst_Princi!CAR_REP_01), "##,##0")
         grd_LisCla.TextMatrix(14, r_int_Col) = ""
         grd_LisCla.TextMatrix(15, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_REP_02), 0, g_rst_Princi!CAR_REP_02), "##,##0")
      End If
      
      r_int_Col = r_int_Col + 1
   Next

End Sub

Private Sub fs_Reporte_Opic(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer, ByVal p_TipExp As Integer)
Dim r_int_Col        As Integer
Dim r_int_Count      As Integer
Dim r_int_MesSel     As Integer

   r_int_Col = 2
   For r_int_Count = 1 To p_PerMes
      If r_int_Count > p_PerMes Then
         Exit For
      End If
      
      r_int_MesSel = r_int_Count
      
      If p_TipExp = 1 Then
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " SELECT "
         g_str_Parame = g_str_Parame & "     ((SELECT ROUND(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP+HIPCIE_SALCON), ((HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM))), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0) "
         g_str_Parame = g_str_Parame & "        + "
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(SUM(DECODE(COMCIE_TIPMON, 1, COMCIE_CAPVIG, (COMCIE_CAPVIG*COMCIE_TIPCAM))), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_COMCIE"
         g_str_Parame = g_str_Parame & "        WHERE COMCIE_PERMES = " & r_int_MesSel & " AND COMCIE_PERANO = " & p_PerAno & ")) AS CAR_VIG_01,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN+HIPCIE_CAPVIG), ((HIPCIE_CAPVEN+HIPCIE_CAPVIG)*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 30 AND HIPCIE_DIAMOR <= 60)   AS CAR_VIG_02,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN+HIPCIE_CAPVIG), ((HIPCIE_CAPVEN+HIPCIE_CAPVIG)*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 60 AND HIPCIE_DIAMOR <= 90)   AS CAR_VIG_03,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN+HIPCIE_CAPVIG), ((HIPCIE_CAPVEN+HIPCIE_CAPVIG)*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 90 AND HIPCIE_DIAMOR <= 120)  AS CAR_VIG_04,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN+HIPCIE_CAPVIG), ((HIPCIE_CAPVEN+HIPCIE_CAPVIG)*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 120 AND HIPCIE_DIAMOR <= 180) AS CAR_VIG_05,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN+HIPCIE_CAPVIG), ((HIPCIE_CAPVEN+HIPCIE_CAPVIG)*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 180 AND HIPCIE_DIAMOR <= 360) AS CAR_VIG_06,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN+HIPCIE_CAPVIG), ((HIPCIE_CAPVEN+HIPCIE_CAPVIG)*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 360)                          AS CAR_VIG_07,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_CAPVEN+HIPCIE_CAPVIG), ((HIPCIE_CAPVEN+HIPCIE_CAPVIG)*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 0 AND HIPCIE_DIAMOR <= 30)    AS CAR_VIG_08,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP+HIPCIE_SALCON), ((HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 1)   AS CAR_REP_01,"
         g_str_Parame = g_str_Parame & "      (SELECT ROUND(NVL(SUM(DECODE(HIPCIE_TIPMON, 1, (HIPCIE_SALCAP+HIPCIE_SALCON), ((HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_TIPCAM))), 0), 2)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 1"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 30)                                                                                AS CAR_REP_02"
         g_str_Parame = g_str_Parame & "  FROM DUAL"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Screen.MousePointer = 0
            Exit Sub
         End If
         
         grd_LisCla.TextMatrix(1, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_01), 0, g_rst_Princi!CAR_VIG_01 - g_rst_Princi!CAR_VIG_02 - g_rst_Princi!CAR_VIG_03 - g_rst_Princi!CAR_VIG_04 - g_rst_Princi!CAR_VIG_05 - g_rst_Princi!CAR_VIG_06 - g_rst_Princi!CAR_VIG_07 - g_rst_Princi!CAR_VIG_08), "###,###,##0.00")
         grd_LisCla.TextMatrix(2, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_02), 0, g_rst_Princi!CAR_VIG_08), "###,###,##0.00")
         grd_LisCla.TextMatrix(3, r_int_Col) = ""
         grd_LisCla.TextMatrix(4, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_02), 0, g_rst_Princi!CAR_VIG_02), "###,###,##0.00")
         grd_LisCla.TextMatrix(5, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_03), 0, g_rst_Princi!CAR_VIG_03), "###,###,##0.00")
         grd_LisCla.TextMatrix(6, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_04), 0, g_rst_Princi!CAR_VIG_04), "###,###,##0.00")
         grd_LisCla.TextMatrix(7, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_05), 0, g_rst_Princi!CAR_VIG_05), "###,###,##0.00")
         grd_LisCla.TextMatrix(8, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_06), 0, g_rst_Princi!CAR_VIG_06), "###,###,##0.00")
         grd_LisCla.TextMatrix(9, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_07), 0, g_rst_Princi!CAR_VIG_07), "###,###,##0.00")
         grd_LisCla.TextMatrix(10, r_int_Col) = ""
         grd_LisCla.TextMatrix(11, r_int_Col) = ""
         grd_LisCla.TextMatrix(12, r_int_Col) = ""
         grd_LisCla.TextMatrix(13, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_REP_01), 0, g_rst_Princi!CAR_REP_01), "###,###,##0.00")
         grd_LisCla.TextMatrix(14, r_int_Col) = ""
         grd_LisCla.TextMatrix(15, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_REP_02), 0, g_rst_Princi!CAR_REP_02), "###,###,##0.00")
      
      Else
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " SELECT "
         g_str_Parame = g_str_Parame & "     ((SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0) "
         g_str_Parame = g_str_Parame & "        + "
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_COMCIE"
         g_str_Parame = g_str_Parame & "        WHERE COMCIE_PERMES = " & r_int_MesSel & " AND COMCIE_PERANO = " & p_PerAno & ")) AS CAR_VIG_01,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 30 AND HIPCIE_DIAMOR <= 60)   AS CAR_VIG_02,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 60 AND HIPCIE_DIAMOR <= 90)   AS CAR_VIG_03,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 90 AND HIPCIE_DIAMOR <= 120)  AS CAR_VIG_04,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 120 AND HIPCIE_DIAMOR <= 180) AS CAR_VIG_05,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 180 AND HIPCIE_DIAMOR <= 360) AS CAR_VIG_06,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 360)                          AS CAR_VIG_07,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 0"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 0 AND HIPCIE_DIAMOR <= 30)    AS CAR_VIG_08,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 1)   AS CAR_REP_01,"
         g_str_Parame = g_str_Parame & "      (SELECT COUNT(*)"
         g_str_Parame = g_str_Parame & "         FROM CRE_HIPCIE"
         g_str_Parame = g_str_Parame & "        WHERE HIPCIE_PERMES = " & r_int_MesSel & " AND HIPCIE_PERANO = " & p_PerAno & " AND HIPCIE_FLGREF = 1"
         g_str_Parame = g_str_Parame & "          AND HIPCIE_DIAMOR > 30)                                                                                AS CAR_REP_02"
         g_str_Parame = g_str_Parame & "  FROM DUAL"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Screen.MousePointer = 0
            Exit Sub
         End If
         
         grd_LisCla.TextMatrix(1, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_01), 0, g_rst_Princi!CAR_VIG_01 - g_rst_Princi!CAR_VIG_02 - g_rst_Princi!CAR_VIG_03 - g_rst_Princi!CAR_VIG_04 - g_rst_Princi!CAR_VIG_05 - g_rst_Princi!CAR_VIG_06 - g_rst_Princi!CAR_VIG_07 - g_rst_Princi!CAR_VIG_08), "##,##0")
         grd_LisCla.TextMatrix(2, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_02), 0, g_rst_Princi!CAR_VIG_08), "##,##0")
         grd_LisCla.TextMatrix(3, r_int_Col) = ""
         grd_LisCla.TextMatrix(4, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_02), 0, g_rst_Princi!CAR_VIG_02), "##,##0")
         grd_LisCla.TextMatrix(5, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_03), 0, g_rst_Princi!CAR_VIG_03), "##,##0")
         grd_LisCla.TextMatrix(6, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_04), 0, g_rst_Princi!CAR_VIG_04), "##,##0")
         grd_LisCla.TextMatrix(7, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_05), 0, g_rst_Princi!CAR_VIG_05), "##,##0")
         grd_LisCla.TextMatrix(8, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_06), 0, g_rst_Princi!CAR_VIG_06), "##,##0")
         grd_LisCla.TextMatrix(9, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_VIG_07), 0, g_rst_Princi!CAR_VIG_07), "##,##0")
         grd_LisCla.TextMatrix(10, r_int_Col) = ""
         grd_LisCla.TextMatrix(11, r_int_Col) = ""
         grd_LisCla.TextMatrix(12, r_int_Col) = ""
         grd_LisCla.TextMatrix(13, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_REP_01), 0, g_rst_Princi!CAR_REP_01), "##,##0")
         grd_LisCla.TextMatrix(14, r_int_Col) = ""
         grd_LisCla.TextMatrix(15, r_int_Col) = Format(IIf(IsNull(g_rst_Princi!CAR_REP_02), 0, g_rst_Princi!CAR_REP_02), "##,##0")
         
      End If
      r_int_Col = r_int_Col + 1
   Next

End Sub

Private Sub fs_GenExcRes()
Dim r_obj_Excel         As Excel.Application
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      'Titulo
      .Cells(1, 1) = "CARTERA EN RIESGO"
      .Range(.Cells(1, 1), .Cells(1, 13)).Merge
      .Range("A1:M1").HorizontalAlignment = xlHAlignCenter
      
      'Primera Linea
      r_int_NroFil = 3
      .Cells(r_int_NroFil, 1) = ""
      .Cells(r_int_NroFil, 2) = "ENERO"
      .Cells(r_int_NroFil, 3) = "FEBRERO"
      .Cells(r_int_NroFil, 4) = "MARZO"
      .Cells(r_int_NroFil, 5) = "ABRIL"
      .Cells(r_int_NroFil, 6) = "MAYO"
      .Cells(r_int_NroFil, 7) = "JUNIO"
      .Cells(r_int_NroFil, 8) = "JULIO"
      .Cells(r_int_NroFil, 9) = "AGOSTO"
      .Cells(r_int_NroFil, 10) = "SETIEMBRE"
      .Cells(r_int_NroFil, 11) = "OCTUBRE"
      .Cells(r_int_NroFil, 12) = "NOVIEMBRE"
      .Cells(r_int_NroFil, 13) = "DICIEMBRE"
      
      .Range(.Cells(1, 1), .Cells(3, 13)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 13)).VerticalAlignment = xlCenter
      .Range(.Cells(3, 1), .Cells(3, 13)).HorizontalAlignment = xlCenter
      .Range(.Cells(3, 1), .Cells(3, 13)).Interior.Color = RGB(146, 208, 80)
      
      'Segunda Linea
      .Columns("A").ColumnWidth = 50:     .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
      .Columns("B").ColumnWidth = 14:     .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 14:     .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 14:     .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 14:     .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 14:     .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 14:     .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 14:     .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 14:     .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 14:     .Cells(r_int_NroFil, 10).HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 14:     .Cells(r_int_NroFil, 11).HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 14:     .Cells(r_int_NroFil, 12).HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 14:     .Cells(r_int_NroFil, 13).HorizontalAlignment = xlHAlignCenter
            
      'Formatea titulo
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 13)).Font.Name = "Calibri"
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 13)).Font.Size = 11
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 13)).Font.Bold = True
      r_int_NroFil = r_int_NroFil + 1
      
      'Exporta filas
      .Range(.Cells(3, 1), .Cells(3, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      
      For r_int_Contad = 4 To 18
         .Cells(r_int_Contad, 1) = grd_LisCla.TextMatrix(r_int_Contad - 3, 1)
         .Cells(r_int_Contad, 2) = grd_LisCla.TextMatrix(r_int_Contad - 3, 2)
         .Cells(r_int_Contad, 3) = grd_LisCla.TextMatrix(r_int_Contad - 3, 3)
         .Cells(r_int_Contad, 4) = grd_LisCla.TextMatrix(r_int_Contad - 3, 4)
         .Cells(r_int_Contad, 5) = grd_LisCla.TextMatrix(r_int_Contad - 3, 5)
         .Cells(r_int_Contad, 6) = grd_LisCla.TextMatrix(r_int_Contad - 3, 6)
         .Cells(r_int_Contad, 7) = grd_LisCla.TextMatrix(r_int_Contad - 3, 7)
         .Cells(r_int_Contad, 8) = grd_LisCla.TextMatrix(r_int_Contad - 3, 8)
         .Cells(r_int_Contad, 9) = grd_LisCla.TextMatrix(r_int_Contad - 3, 9)
         .Cells(r_int_Contad, 10) = grd_LisCla.TextMatrix(r_int_Contad - 3, 10)
         .Cells(r_int_Contad, 11) = grd_LisCla.TextMatrix(r_int_Contad - 3, 11)
         .Cells(r_int_Contad, 12) = grd_LisCla.TextMatrix(r_int_Contad - 3, 12)
         .Cells(r_int_Contad, 13) = grd_LisCla.TextMatrix(r_int_Contad - 3, 13)
         .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      Next
      
      .Range(.Cells(3, 1), .Cells(r_int_Contad - 1, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 2), .Cells(r_int_Contad - 1, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 3), .Cells(r_int_Contad - 1, 3)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 4), .Cells(r_int_Contad - 1, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 5), .Cells(r_int_Contad - 1, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 6), .Cells(r_int_Contad - 1, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 7), .Cells(r_int_Contad - 1, 7)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 8), .Cells(r_int_Contad - 1, 8)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 9), .Cells(r_int_Contad - 1, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 10), .Cells(r_int_Contad - 1, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 11), .Cells(r_int_Contad - 1, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 12), .Cells(r_int_Contad - 1, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 13), .Cells(r_int_Contad - 1, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(3, 14), .Cells(r_int_Contad - 1, 14)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      
      .Range(.Cells(r_int_Contad, 1), .Cells(r_int_Contad, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipRep)
   End If
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipRep.ListIndex > -1 Then
         Call gs_SetFocus(cmb_TipDat)
      End If
   End If
End Sub

Private Sub cmb_TipDat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipDat.ListIndex > -1 Then
         Call gs_SetFocus(cmd_Proces)
      End If
   End If
End Sub

