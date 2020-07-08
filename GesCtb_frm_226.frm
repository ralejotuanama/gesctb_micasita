VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptCtb_34 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10350
   Icon            =   "GesCtb_frm_226.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6285
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10365
      _Version        =   65536
      _ExtentX        =   18283
      _ExtentY        =   11086
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   855
         Left            =   60
         TabIndex        =   7
         Top             =   2160
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
         _ExtentY        =   1508
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
         Begin VB.ComboBox cmb_CodPro 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   6255
         End
         Begin VB.Label Label4 
            Caption         =   "Banco:"
            Height          =   375
            Left            =   150
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
         _ExtentY        =   1244
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
            Height          =   585
            Left            =   630
            TabIndex        =   9
            Top             =   90
            Width           =   8865
            _Version        =   65536
            _ExtentX        =   15637
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Procesos - Carga Archivos Contabilidad"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            Picture         =   "GesCtb_frm_226.frx":000C
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   60
         TabIndex        =   10
         Top             =   810
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
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
            Left            =   9630
            Picture         =   "GesCtb_frm_226.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   60
            Picture         =   "GesCtb_frm_226.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Cargar Archivo"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2295
         Left            =   60
         TabIndex        =   12
         Top             =   3060
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
         _ExtentY        =   4048
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
         Begin VB.DriveListBox drv_LisUni 
            Height          =   315
            Left            =   6090
            TabIndex        =   18
            Top             =   60
            Width           =   4095
         End
         Begin VB.DirListBox dir_LisCar 
            Height          =   1665
            Left            =   6075
            TabIndex        =   2
            Top             =   420
            Width           =   4095
         End
         Begin VB.FileListBox fil_LisArc 
            Height          =   2040
            Left            =   1590
            TabIndex        =   3
            Top             =   90
            Width           =   4425
         End
         Begin VB.Label Label1 
            Caption         =   "Archivo a cargar:"
            Height          =   315
            Left            =   150
            TabIndex        =   13
            Top             =   90
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   825
         Left            =   60
         TabIndex        =   14
         Top             =   5400
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
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
         Begin Threed.SSPanel pnl_BarPro 
            Height          =   345
            Left            =   60
            TabIndex        =   19
            Top             =   420
            Width           =   10155
            _Version        =   65536
            _ExtentX        =   17912
            _ExtentY        =   609
            _StockProps     =   15
            Caption         =   "SSPanel2"
            ForeColor       =   16777215
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
            FloodType       =   1
            FloodColor      =   49152
            Font3D          =   2
         End
         Begin VB.Label lbl_NomPro 
            Caption         =   "Proceso carga información..."
            Height          =   255
            Left            =   90
            TabIndex        =   15
            Top             =   120
            Width           =   5505
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   645
         Left            =   60
         TabIndex        =   16
         Top             =   1500
         Width           =   10245
         _Version        =   65536
         _ExtentX        =   18071
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
         Begin VB.ComboBox cmb_TipCar 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   6255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Carga:"
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   240
            Width           =   1050
         End
      End
   End
End
Attribute VB_Name = "frm_RptCtb_34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Bancos()   As moddat_tpo_Genera
Dim l_bol_estado As Boolean
Private Type r_Arr_CargaExtractBancos
    r_str_Bank          As String
    r_str_FecExt        As String
    r_str_Ope           As String
    r_str_Concept       As String
    r_dbl_Imp           As Double
    r_str_CtaBank       As String
End Type

Dim r_str_Bancos()      As r_Arr_CargaExtractBancos
Private Type r_Arr_Calendario
   NumOpe               As String
   OpeMvi               As String
   Secuen               As String
   NroCuo               As String
   FecVct               As String
   NroDia               As String
   Moneda               As String
   TipCro               As Integer     '3 y 4
   Capita               As Double
   Intere               As Double
   ComCof               As Double
   MtoCuo               As Double
   SalCap               As Double
End Type

Dim r_str_CalDes()      As r_Arr_Calendario
Private Type r_Arr_DesErr
   NomArc               As String
   NumOpe               As String
   OpeMvi               As String
   MtoErr               As String
End Type
Dim r_str_DesErr()      As r_Arr_DesErr
Private Sub cmb_TipCar_Click()
   If cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 1 Then
      cmb_CodPro.Enabled = True
      lbl_NomPro.Caption = "Proceso carga Extracto Bancario"
      fil_LisArc.Pattern = "*.xls"
   End If
End Sub
Private Sub cmb_TipCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodPro)
   End If
End Sub

Private Sub cmd_Proces_Click()
Dim r_lng_Contad           As Long
Dim r_str_NomArc           As String
Dim r_lng_NumReg           As Long
Dim r_lng_TotReg           As Long
Dim modprc_g_str_CadEje    As String
Dim r_str_ctabco           As String

   If cmb_TipCar.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Carga.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCar)
      Exit Sub
   End If

   If cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 1 Then
      If Me.cmb_CodPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar un Banco.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CodPro)
         Exit Sub
      End If
      If Len(Trim(fil_LisArc.FileName & "")) = 0 Then
         MsgBox "Debe seleccionar el Archivo a cargar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de ejecutar el proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   cmd_Proces.Enabled = False
   
   r_str_ctabco = Left(fil_LisArc.FileName, 12)
   
   Call fs_ValidarCargaExtractoBanco(fil_LisArc.Path & "\" & fil_LisArc.FileName, l_arr_Bancos(cmb_CodPro.ListIndex + 1).Genera_Codigo, r_lng_TotReg, r_str_ctabco)
   
   If cmb_TipCar.ItemData(cmb_TipCar.ListIndex) = 1 Then               'Carga Saldos COFIDE
       If (r_lng_TotReg > -1) Then
           lbl_NomPro.Caption = "Proceso carga Extracto Bancario...": DoEvents
           pnl_BarPro.FloodPercent = 0
           
           For r_lng_Contad = 1 To UBound(r_str_Bancos)
        
              r_lng_NumReg = r_lng_Contad
              'Inserta
              g_str_Parame = ""
              g_str_Parame = "INSERT INTO CTB_CONBAN ("
              g_str_Parame = g_str_Parame & "CONBAN_CODBCO, "
              g_str_Parame = g_str_Parame & "CONBAN_FECMOV, "
              g_str_Parame = g_str_Parame & "CONBAN_NUMOPE, "
              g_str_Parame = g_str_Parame & "CONBAN_CONCEP, "
              g_str_Parame = g_str_Parame & "CONBAN_IMPORT, "
              g_str_Parame = g_str_Parame & "CONBAN_CTABCO) "
              g_str_Parame = g_str_Parame & "VALUES ( "
              g_str_Parame = g_str_Parame & "'" & r_str_Bancos(r_lng_Contad).r_str_Bank & "', "
              g_str_Parame = g_str_Parame & r_str_Bancos(r_lng_Contad).r_str_FecExt & ", "
              g_str_Parame = g_str_Parame & "'" & r_str_Bancos(r_lng_Contad).r_str_Ope & "', "
              g_str_Parame = g_str_Parame & "'" & r_str_Bancos(r_lng_Contad).r_str_Concept & "', "
              g_str_Parame = g_str_Parame & r_str_Bancos(r_lng_Contad).r_dbl_Imp & ", "
              g_str_Parame = g_str_Parame & "'" & r_str_Bancos(r_lng_Contad).r_str_CtaBank & "')"
             
              If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  Exit Sub
              End If
              DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
             
              pnl_BarPro.FloodPercent = CDbl(Format(r_lng_NumReg / r_lng_TotReg * 100, "##0.00"))
             
           Next r_lng_Contad
      End If
      cmd_Proces.Enabled = True
      Screen.MousePointer = 0

      MsgBox "Proceso Terminado.", vbInformation, modgen_g_str_NomPlt
      
   End If
End Sub
Private Sub fs_ValidarCargaExtractoBanco(ByVal filePath As String, ByVal bank As String, ByRef r_lng_TotRegNuevos As Long, ByVal r_str_ctabco As String)
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_str_FecExt        As String
Dim r_str_Ope           As String
Dim r_str_Concept       As String
Dim r_str_Aux           As String
Dim r_dbl_Imp           As Double

Dim r_int_Judici        As String
Dim r_int_Filaux        As Integer
Dim r_str_NumCon        As String
Dim r_str_NumCoC        As String
Dim r_str_TIPMON        As String

Dim r_lng_Contad        As Long
Dim r_lng_NumReg        As Long

   'BCP
   If (StrComp(bank, "000001") = 0) And ((StrComp(r_str_ctabco, "111301030101") <> 0) And (StrComp(r_str_ctabco, "111301030101") <> 0)) Then
      MsgBox "El nombre del archivo a cargar debe empezar con el número de cuenta del banco BCP", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_lng_TotRegNuevos = -1
      Exit Sub
   'Interbank
   ElseIf (bank = "000004") And (r_str_ctabco <> "111301040201") Then
      MsgBox "El nombre del archivo a cargar debe empezar con el número de cuenta del banco Interbank", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_lng_TotRegNuevos = -1
      Exit Sub
   'BBVA
   ElseIf (bank = "000002") And ((r_str_ctabco <> "111301060102") And (r_str_ctabco <> "111301060103") And (r_str_ctabco <> "111301060201") And (r_str_ctabco <> "112301060102") And (r_str_ctabco <> "112301060202")) Then
      MsgBox "El nombre del archivo a cargar debe empezar con el número de cuenta del banco BBVA", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_lng_TotRegNuevos = -1
      Exit Sub
   'GNB
   ElseIf (bank = "000009") And ((r_str_ctabco <> "111301320101") And (r_str_ctabco <> "112301320101")) Then
      MsgBox "El nombre del archivo a cargar debe empezar con el número de cuenta del banco GNB", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_lng_TotRegNuevos = -1
      Exit Sub
   'BANBIF
   ElseIf (bank = "000005") And (r_str_ctabco <> "111301210101") Then
      MsgBox "El nombre del archivo a cargar debe empezar con el número de cuenta del banco BANBIF", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_lng_TotRegNuevos = -1
      Exit Sub
   ElseIf (bank <> "000001") And (bank <> "000004") And (bank <> "000002") And (bank <> "000009") And (bank <> "000005") Then
      MsgBox "Banco con cuentas sin configurar", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_lng_TotRegNuevos = -1
      Exit Sub
   End If
   
    'Abriendo Archivo Extracto Banco
   Set r_obj_Excel = New Excel.Application
   
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=filePath
   r_int_FilExc = 1
   r_int_Filaux = 0
   r_int_Judici = 0
   
   r_lng_NumReg = 0
   r_lng_TotRegNuevos = 0
   
   lbl_NomPro.Caption = "Validando el Archivo " + filePath + " ...": DoEvents
   
   ReDim r_str_Bancos(0)
   
   r_obj_Excel.Sheets(1).Select
   
   r_str_Ope = Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value)
   If (r_str_Ope <> "FECHA") Then
      MsgBox "La primera cabecera debe llamarse: FECHA", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_lng_TotRegNuevos = -1
      Exit Sub
   End If
   r_str_Ope = Trim(r_obj_Excel.Cells(r_int_FilExc, 2).Value)
   If (r_str_Ope <> "OPERACION") Then
      MsgBox "La segunda cabecera debe llamarse: OPERACION", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_lng_TotRegNuevos = -1
      Exit Sub
   End If
   r_str_Concept = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
   If (r_str_Concept <> "CONCEPTO") Then
      MsgBox "La tercera cabecera debe llamarse: CONCEPTO", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_lng_TotRegNuevos = -1
      Exit Sub
   End If
   r_str_Concept = Trim(r_obj_Excel.Cells(r_int_FilExc, 4).Value)
   If (r_str_Concept <> "IMPORTE") Then
      MsgBox "La cuarta cabecera debe llamarse: IMPORTE", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_lng_TotRegNuevos = -1
      Exit Sub
   End If
   
   r_int_FilExc = r_int_FilExc + 1
   
   'HOJA PRINCIPAL
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value) <> ""
      r_str_FecExt = Trim(r_obj_Excel.Cells(r_int_FilExc, 1).Value)
      If IsNumeric(r_str_FecExt) Then
         MsgBox "El valor de la columna FECHA en la celda A" + CStr(r_int_FilExc) + " debe ser una fecha con formato dd/mm/yyyy.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      ElseIf (Len(r_str_FecExt) <> 10) Then
         MsgBox "El valor de la columna FECHA en la celda A" + CStr(r_int_FilExc) + " debe ser de máximo 10 dígitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      ElseIf (InStr(r_str_FecExt, "/") = 0) Then
         MsgBox "El valor de la columna FECHA en la celda A" + CStr(r_int_FilExc) + " debe ser una fecha con formato dd/mm/yyyy.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      ElseIf Not (IsNumeric(Left(r_str_FecExt, 2)) And IsNumeric(Mid(r_str_FecExt, 4, 2)) And IsNumeric(Right(r_str_FecExt, 4))) Then
         MsgBox "El valor de la columna FECHA en la celda A" + CStr(r_int_FilExc) + " debe ser una fecha con formato dd/mm/yyyy donde dd, mm y yyyy deben ser valores númericos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      ElseIf Not ((CInt(Left(r_str_FecExt, 2)) > 0) And (CInt(Left(r_str_FecExt, 2)) < 32) And (CInt(Mid(r_str_FecExt, 4, 2)) > 0) And (CInt(Mid(r_str_FecExt, 4, 2)) < 13) And (CInt(Right(r_str_FecExt, 4)) > 1899) And (CInt(Right(r_str_FecExt, 4)) < 3000)) Then
         MsgBox "El valor de la columna FECHA en la celda A" + CStr(r_int_FilExc) + " debe ser una fecha con formato dd/mm/yyyy con fechas correctas.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      End If
      r_str_Ope = Trim(r_obj_Excel.Cells(r_int_FilExc, 2).Value)
      If Not IsNumeric(r_str_Ope) Then
         MsgBox "El valor de la columna OPERACION en la celda B" + CStr(r_int_FilExc) + " debe ser un valor numérico entero.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      ElseIf (Len(r_str_Ope) > 10) Then
         MsgBox "El valor de la columna OPERACION en la celda B" + CStr(r_int_FilExc) + " debe ser de máximo 10 dígitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      ElseIf (InStr(r_str_Ope, ".") > 0) Then
         MsgBox "El valor de la columna OPERACION en la celda B" + CStr(r_int_FilExc) + " debe ser un valor numérico entero.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      End If
      r_str_Concept = Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value)
      If (IsNull(r_str_Concept) Or (r_str_Concept = "")) Then
         MsgBox "El valor de la columna CONCEPTO en la celda C" + CStr(r_int_FilExc) + " no puede ser vacío.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      ElseIf IsNumeric(r_str_Concept) Then
         MsgBox "El valor de la columna CONCEPTO en la celda C" + CStr(r_int_FilExc) + " debe ser alfanumérico.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      End If
      r_str_Aux = Trim(r_obj_Excel.Cells(r_int_FilExc, 4).Value)
      If Not IsNumeric(r_str_Aux) Then
         MsgBox "El valor de la columna IMPORTE en la celda D" + CStr(r_int_FilExc) + " debe ser un valor numérico no vacío.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      End If
      
      g_str_Parame = "SELECT NVL(COUNT(*),0) AS TOTREG FROM CTB_CONBAN WHERE CONBAN_CODBCO = " & bank & " AND CONBAN_NUMOPE = " & r_str_Ope
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "Problemas de conexión con base de datos, intente la carga más tarde.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
      End If
      
      g_rst_Princi.MoveFirst
      r_lng_Contad = g_rst_Princi!TOTREG
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      If (r_lng_Contad = 0) Then
        'Cargar la data en el array
        ReDim Preserve r_str_Bancos(UBound(r_str_Bancos) + 1)
        r_str_Bancos(UBound(r_str_Bancos)).r_str_Bank = bank
        r_str_Bancos(UBound(r_str_Bancos)).r_str_FecExt = Format(r_str_FecExt, "yyyymmdd")
        r_str_Bancos(UBound(r_str_Bancos)).r_str_Ope = r_str_Ope
        r_str_Bancos(UBound(r_str_Bancos)).r_str_Concept = r_str_Concept
        r_str_Bancos(UBound(r_str_Bancos)).r_dbl_Imp = CDbl(Val(r_str_Aux))
        r_str_Bancos(UBound(r_str_Bancos)).r_str_CtaBank = r_str_ctabco
        r_lng_TotRegNuevos = r_lng_TotRegNuevos + 1
      End If
      r_int_FilExc = r_int_FilExc + 1
      r_lng_NumReg = r_lng_NumReg + 1

   Loop
   
   If (Trim(r_obj_Excel.Cells(r_int_FilExc, 2).Value) <> "") Or (Trim(r_obj_Excel.Cells(r_int_FilExc, 3).Value) <> "") Or (Trim(r_obj_Excel.Cells(r_int_FilExc, 4).Value) <> "") Then
         MsgBox "El valor de la columna FECHA en la celda A" + CStr(r_int_FilExc) + " debe ser una fecha con formato dd/mm/yyyy no vacío.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(fil_LisArc)
         r_lng_TotRegNuevos = -1
         Exit Sub
   End If
   
   r_obj_Excel.Workbooks.Close
   Set r_obj_Excel = Nothing
   
   If (r_lng_TotRegNuevos = 0) Then
      MsgBox "El archivo no posee nuevas operaciones para el banco seleccionado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_obj_Excel.Workbooks.Close
      Exit Sub
   End If
   
   If (r_lng_NumReg = 0) Then
      MsgBox "El archivo no posee registros despues de la cabecera.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(fil_LisArc)
      r_obj_Excel.Workbooks.Close
      Exit Sub
   End If
   
End Sub
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub drv_LisUni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error Resume Next
      dir_LisCar.Path = drv_LisUni.Drive
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   
   cmb_TipCar.Clear
   cmb_TipCar.AddItem "EXTRACTO BANCARIO"
   cmb_TipCar.ItemData(cmb_TipCar.NewIndex) = CInt(1)
   cmb_TipCar.ListIndex = -1
   
   Call moddat_gs_Carga_LisIte(cmb_CodPro, l_arr_Bancos, 1, "513")
   cmb_CodPro.ListIndex = -1
   
   dir_LisCar.Path = "C:\"
   lbl_NomPro.Caption = Empty
End Sub

Private Sub fs_Limpia()
   fil_LisArc.Pattern = "*.xls"
End Sub

Private Sub cmb_CodPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(drv_LisUni)
   End If
End Sub

Private Sub drv_LisUni_Change()
   On Error Resume Next
   dir_LisCar.Path = drv_LisUni.Drive
End Sub

Private Sub dir_LisCar_Change()
   fil_LisArc.Path = dir_LisCar.Path
End Sub

