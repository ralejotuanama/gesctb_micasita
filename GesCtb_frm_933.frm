VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form Frm_Ctb_FacEle_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14190
   Icon            =   "GesCtb_frm_933.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14205
      _Version        =   65536
      _ExtentX        =   25056
      _ExtentY        =   17489
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   45
         TabIndex        =   1
         Top             =   45
         Width           =   14115
         _Version        =   65536
         _ExtentX        =   24897
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   300
            Left            =   660
            TabIndex        =   2
            Top             =   180
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Registro de Documentos Electrónicos"
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
         Begin VB.Image Image2 
            Height          =   480
            Left            =   60
            Picture         =   "GesCtb_frm_933.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   45
         TabIndex        =   3
         Top             =   750
         Width           =   14115
         _Version        =   65536
         _ExtentX        =   24897
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
            Left            =   12360
            Picture         =   "GesCtb_frm_933.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Grabar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   12945
            Picture         =   "GesCtb_frm_933.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Cancelar "
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   13515
            Picture         =   "GesCtb_frm_933.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   240
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin Threed.SSPanel SSPanel7 
         Height          =   7725
         Left            =   30
         TabIndex        =   7
         Top             =   2130
         Width           =   14115
         _Version        =   65536
         _ExtentX        =   24897
         _ExtentY        =   13626
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
            Height          =   7605
            Left            =   30
            TabIndex        =   8
            Top             =   60
            Width           =   13995
            _ExtentX        =   24686
            _ExtentY        =   13414
            _Version        =   393216
            Rows            =   30
            Cols            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
            Appearance      =   0
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   1440
         Width           =   14115
         _Version        =   65536
         _ExtentX        =   24897
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
         Begin VB.TextBox txt_NomArc 
            Height          =   315
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "txt_NomArc"
            Top             =   180
            Width           =   10515
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
            Left            =   12810
            TabIndex        =   11
            ToolTipText     =   "Seleccionar archivo"
            Top             =   180
            Width           =   315
         End
         Begin VB.CommandButton cmd_Import 
            Height          =   585
            Left            =   13485
            Picture         =   "GesCtb_frm_933.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Importar archivo"
            Top             =   30
            Width           =   585
         End
         Begin VB.Label Label3 
            Caption         =   "Archivo a cargar:"
            Height          =   255
            Left            =   300
            TabIndex        =   13
            Top             =   210
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "Frm_Ctb_FacEle_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_RegTot        As Integer
Dim l_int_TipOpc        As Integer
Dim l_int_NumReg        As Integer
Dim l_bol_FlgErr        As Boolean
Dim l_int_RegPro        As Integer
Dim l_int_RegErr        As Integer
Dim l_int_RegSPr        As Integer
Dim l_lng_Codigo        As Long
Dim l_str_MsjRef        As String
Dim l_str_RutaLg        As String
Dim l_fsobj             As Scripting.FileSystemObject
Dim l_txtStr            As TextStream

Private Type g_tpo_CarDocEle
   CarDocEle_Col1       As String      'NumRef
   CarDocEle_ColErr1    As String
   CarDocEle_Col2       As String      'TipCom
   CarDocEle_ColErr2    As String
   CarDocEle_Col3       As String      'TipPro
   CarDocEle_ColErr3    As String
   CarDocEle_Col4       As String      'FecEmi
   CarDocEle_ColErr4    As String
   CarDocEle_Col5       As String      'Moneda
   CarDocEle_ColErr5    As String
   CarDocEle_Col6       As String      'TipCam
   CarDocEle_ColErr6    As String
   CarDocEle_Col7       As String      'TipDoc
   CarDocEle_ColErr7    As String
   CarDocEle_Col8       As String      'NumDoc
   CarDocEle_ColErr8    As String
   CarDocEle_Col9       As String      'Direcc
   CarDocEle_ColErr9    As String
   CarDocEle_Col10      As String      'Distri
   CarDocEle_ColErr10   As String
   CarDocEle_Col11      As String      'Provin
   CarDocEle_ColErr11   As String
   CarDocEle_Col12      As String      'Depart
   CarDocEle_ColErr12   As String
   CarDocEle_Col13      As String      'Correo
   CarDocEle_ColErr13   As String
   CarDocEle_Col14      As String      'Cantid
   CarDocEle_ColErr14   As String
   CarDocEle_Col15      As String      'Codigo
   CarDocEle_ColErr15   As String
   CarDocEle_Col16      As String      'UniMed
   CarDocEle_ColErr16   As String
   CarDocEle_Col17      As String      'Glosa
   CarDocEle_ColErr17   As String
   CarDocEle_Col18      As String      'ValUni
   CarDocEle_ColErr18   As String
   CarDocEle_Col19      As String      'VtaTot
   CarDocEle_ColErr19   As String
   CarDocEle_Col20      As String      'RazSoc
   CarDocEle_ColErr20   As String
   CarDocEle_Col21      As String      'Observación
   CarDocEle_ColErr21   As String
End Type

Private Sub cmd_BuscaArc_Click()
   dlg_Guarda.Filter = "Archivos Excel |*.xlsx;*.xls"
   dlg_Guarda.ShowOpen
   txt_NomArc.Text = UCase(dlg_Guarda.FileName)
   Exit Sub
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpiar
   Call fs_Activa(False)
End Sub

Private Sub cmd_Grabar_Click()
   If grd_Listad.Rows = 0 Then Exit Sub
      
   If l_bol_FlgErr = False Then
      If MsgBox("¿Desea cargar a la base de datos la información de los documentos electónicos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         l_int_TipOpc = 0
         Call fs_Inicia
         Call fs_Limpiar
         Call fs_Activa(False)
         Exit Sub
      End If
      
      'proceso de grabación
      Screen.MousePointer = 11
      If fs_Agrega_Grilla_DocEle() Then
         MsgBox "Proceso realizado satisfactoriamente ", vbInformation, modgen_g_str_NomPlt
      Else
         MsgBox "Registros Encontrados" & vbTab & ":" & "   " & l_int_RegTot & vbNewLine & _
                "Registros Procesados" & vbTab & ":" & "   " & l_int_RegPro & vbNewLine & _
                "Registros Errados" & vbTab & vbTab & ":" & "   " & l_int_RegErr & vbNewLine & _
                "Registros Sin Procesar" & vbTab & ":" & "   " & l_int_RegSPr, vbInformation, modgen_g_str_NomPlt
      End If
      Screen.MousePointer = 0
   Else
       MsgBox "No se puede cargar la información a la base de datos, aún existen errores.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(grd_Listad)
       Exit Sub
   End If
   
   Call fs_Inicia
   Call fs_Activa(False)
   Call Frm_Ctb_FacEle_01.fs_Buscar
End Sub

Private Sub cmd_Import_Click()
   l_int_TipOpc = 1
   'validaciones
   If Len(Trim(txt_NomArc.Text)) = 0 Then
      MsgBox "Debe ingresar la ubicación y nombre del archivo a importar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If
   
   MsgBox "Antes de realizar el proceso de carga verifique los siguiente: " & vbCrLf & " - El archivo excel debe tener el formato minimo del 2007. " & vbCrLf & " - La Columna D del archivo con formato 'dd/mm/yyyy'", vbInformation, modgen_g_str_NomPlt
   If MsgBox("¿Desea realizar la carga del archivo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   If fs_Cargar_ArcDocEle Then
      If l_int_NumReg = 0 Then
         MsgBox "No se encontraron datos para importar.", vbInformation, modgen_g_str_NomPlt
      Else
         If l_bol_FlgErr = True Then
            MsgBox "Se encontraron errores al importar el archivo.", vbInformation, modgen_g_str_NomPlt
         Else
            MsgBox "Proceso de carga de archivos, finalizado satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
            cmd_Grabar.Enabled = True
         End If
      End If
   Else
      MsgBox "El archivo no cumple con el formato preestablecido.", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpiar()
   Call gs_LimpiaGrid(grd_Listad)
   Screen.MousePointer = 0
   txt_NomArc.Text = ""
End Sub

Private Function fs_Cargar_ArcDocEle() As Boolean
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_int_Contad        As Integer
Dim r_int_TotReg        As Integer
Dim r_int_ConAux        As Integer

'Cabecera
Dim r_str_PerAnn        As String
Dim r_str_PerMes        As String
Dim r_lng_NumRef        As Long
Dim r_str_Fecemi        As String
Dim r_str_TipCom        As String
Dim r_str_TipPro        As String
Dim r_int_TipPro        As Integer
Dim r_str_TIPMON        As String
Dim r_dbl_TipCam        As Double
Dim r_str_GloDet        As String
Dim r_str_TipDoc        As String
Dim r_str_NumDoc        As String

'Detalle
Dim r_int_Cantid        As Integer
Dim r_str_Codigo        As String
Dim r_str_UniMed        As String
Dim r_dbl_MtoUni        As Double
Dim r_dbl_MtoVta        As Double
Dim r_arr_Matriz()      As g_tpo_CarDocEle
Dim r_str_Recept        As String
Dim r_str_Direcc        As String
Dim r_str_Distri        As String
Dim r_str_Provin        As String
Dim r_str_Depart        As String
Dim r_str_Correo        As String
Dim r_str_Direcc_bas    As String
Dim r_str_Distri_bas    As String
Dim r_str_Provin_bas    As String
Dim r_str_Depart_bas    As String
Dim r_str_Correo_bas    As String
Dim r_str_Observ        As String
   
   l_bol_FlgErr = False
   fs_Cargar_ArcDocEle = False
   Call gs_LimpiaGrid(grd_Listad)
   ReDim r_arr_Matriz(0)
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=txt_NomArc.Text
   
   r_int_FilExc = 0
   If fs_Verificar_Cabecera(r_obj_Excel) = True Then
          
      Do While Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value) <> ""
         ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
         
         'Número de Referencia
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)) Then
            r_lng_NumRef = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col1 = r_lng_NumRef
         Else
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col1 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr1 = "Número de Referencia Inválido."
         End If
         
         'Tipo de Comprobante
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value)) And (Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) = "F" Or Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) = "B") Then
            r_str_TipCom = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col2 = r_str_TipCom
         Else
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col2 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr2 = "Tipo de Comprobante Inválido."
         End If
         
         'Valida Tipo Proceso
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)) And fs_Validar_TipPro(Mid(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value), 1, 1)) = True Then
             r_str_TipPro = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
             r_int_TipPro = Mid(r_str_TipPro, 1, 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col3 = r_str_TipPro
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col3 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr3 = "Tipo de Operación Inválida."
         End If
                    
         'Valida Fecha de Emisión
         If IsDate(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)) Then
            'And Year(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)) = r_str_PerAnn And
            'Month(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)) = r_str_PerMes
             r_str_Fecemi = Format(Year(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)), "0000") & "-" & Format(Month(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)), "00") & "-" & Format(Day(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)), "00")
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col4 = r_str_Fecemi
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col4 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr4 = "Fecha Inválida."
         End If

         'Valida Moneda
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value)) And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value) <> "" Then
             r_str_TIPMON = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col5 = r_str_TIPMON
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col5 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr5 = "Tipo de Moneda Inválida."
         End If

         'Valida Tipo de Cambio
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value)) And r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value <> "" Then
            If r_str_TIPMON = "SOLES" And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value) = 0 Then
               r_dbl_TipCam = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value)
               r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col6 = r_dbl_TipCam
            Else
               r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col6 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value)
               r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr6 = "Tipo de Cambio Inválido."
            End If
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col6 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr6 = "Tipo de Cambio Inválido."
         End If
          
         'Valida TipoDoc
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value)) And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value) <> "" Then
             r_str_TipDoc = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col7 = r_str_TipDoc
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col7 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr7 = "Tipo de Documento inválido."
         End If

         'Valida NumDoc
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value)) And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value) <> "" And fs_Validar_NumDoc(Mid(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value), 1, 1), Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value)) = True Then
             r_str_NumDoc = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col8 = r_str_NumDoc
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col8 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr8 = "Número de Documento inválido."
         End If
         
         'Valida Razón Social
         If Not IsNull(r_str_NumDoc) And Not IsNull(r_str_NumDoc) Then
             Call fs_Buscar_Receptor(r_int_TipPro, CInt(Mid(r_str_TipDoc, 1, 1)), r_str_NumDoc, r_str_Recept, r_str_Direcc_bas, r_str_Distri_bas, r_str_Provin_bas, r_str_Depart_bas, r_str_Correo_bas)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col20 = r_str_Recept
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col20 = ""
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr20 = "No se encontró el Receptor."
         End If
         
         'Valida Direccion
         If r_str_Direcc_bas = "" Or IsNull(r_str_Direcc_bas) Then
            r_str_Direcc = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col9 = r_str_Direcc
         Else
            r_str_Direcc = r_str_Direcc_bas
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col9 = r_str_Direcc
         End If
             
         'Valida Distrito
         If r_str_Distri_bas = "" Or IsNull(r_str_Distri_bas) Then
            r_str_Distri = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col10 = r_str_Distri
         Else
            r_str_Distri = r_str_Distri_bas
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col10 = r_str_Distri
         End If
         
         'Valida Provincia
         If r_str_Provin_bas = "" Or IsNull(r_str_Provin_bas) Then
            r_str_Provin = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 11).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col11 = r_str_Provin
         Else
            r_str_Provin = r_str_Provin_bas
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col11 = r_str_Provin
         End If
         
         'Valida Departamento
         If r_str_Depart_bas = "" Or IsNull(r_str_Depart_bas) Then
            r_str_Depart = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 12).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col12 = r_str_Depart
         Else
            r_str_Depart = r_str_Depart_bas
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col12 = r_str_Depart
         End If
         
         'Valida Correo
         If r_str_Correo_bas = "" Or IsNull(r_str_Correo_bas) Then
            r_str_Correo = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 13).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col13 = r_str_Correo
         Else
            r_str_Correo = r_str_Correo_bas
            r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col13 = r_str_Correo
         End If
         
         'Valida Cantidad
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 14).Value)) And IsNumeric(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 14).Value)) Then
             r_int_Cantid = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 14).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col14 = r_int_Cantid
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col14 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 14).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr14 = "No tiene Cantidad."
         End If
            
         'Valida Código
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 15).Value)) And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 15).Value) <> "" Then
             r_str_Codigo = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 15).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col15 = r_str_Codigo
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col15 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 15).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr15 = "Fecha Inválida."
         End If
            
         'Valida Unidad de Medida
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 16).Value)) And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 16).Value) <> "" Then
             r_str_UniMed = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 16).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col16 = r_str_UniMed
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col16 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 16).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr16 = "No tiene Unidad de Medida."
         End If
         
         'Valida Glosa
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 17).Value)) And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 17).Value) <> "" Then
             r_str_GloDet = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 17).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col17 = r_str_GloDet
         Else
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col17 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 17).Value)
             r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr17 = "No tiene Glosa."
         End If
            
         'Valida Valor Unitario
         If IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 18).Value)) Or Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 18).Value) = "" Then
             r_dbl_MtoUni = 0
         Else
             r_dbl_MtoUni = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 18).Value)
         End If
         r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col18 = r_dbl_MtoUni
            
         'Valida Valor Total Venta
         If IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 19).Value)) Or Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 19).Value) = "" Then
             r_dbl_MtoVta = 0
         Else
             r_dbl_MtoVta = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 19).Value)
         End If
         r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col19 = Format(r_dbl_MtoVta, "###,###,##0.00")
         
         'Valida Observación
         If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 20).Value)) And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 20).Value) <> "" Then
            r_str_Observ = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 20).Value)
            'Validamos si existe la referencia
            If fs_Validar_Referencia(r_str_Observ) = True Then
               r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col21 = r_str_Observ
            Else
               If r_int_TipPro = 3 Or r_int_TipPro = 4 Or r_int_TipPro = 6 Then
                  r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col21 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 20).Value)
                  r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr21 = "No existe la Referencia."
               End If
            End If
         Else
            If r_int_TipPro = 3 Or r_int_TipPro = 4 Or r_int_TipPro = 6 Then
               r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col21 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 20).Value)
               r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_ColErr21 = "No tiene Referencia."
            End If
         End If
         
         r_int_FilExc = r_int_FilExc + 1
      Loop
       
      l_int_NumReg = r_int_FilExc
      grd_Listad.Redraw = False
      
      'Cabecera de la Grilla
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.FixedRows = 1
      grd_Listad.Rows = grd_Listad.Rows - 1
      grd_Listad.Row = 0
      
      grd_Listad.Col = 0:        grd_Listad.Text = "NUMREF":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 1:        grd_Listad.Text = "TIPCOM":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 2:        grd_Listad.Text = "TIPPRO":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 3:        grd_Listad.Text = "FECEMI":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 4:        grd_Listad.Text = "MONEDA":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 5:        grd_Listad.Text = "TIPCAM":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 6:        grd_Listad.Text = "TIPDOC":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 7:        grd_Listad.Text = "NUMDOC":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 8:        grd_Listad.Text = "RECEPTOR":  grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 9:        grd_Listad.Text = "DIRECCION": grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 10:       grd_Listad.Text = "DISTRITO":  grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 11:       grd_Listad.Text = "PROVINCIA": grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 12:       grd_Listad.Text = "DEPARTAMENTO": grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 13:       grd_Listad.Text = "CORREO":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 14:       grd_Listad.Text = "CANTIDAD":  grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 15:       grd_Listad.Text = "CODIGO":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 16:       grd_Listad.Text = "UM":        grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 17:       grd_Listad.Text = "GLOSA":     grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 18:       grd_Listad.Text = "VALUNI":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 19:       grd_Listad.Text = "VALVTA":    grd_Listad.CellAlignment = flexAlignCenterCenter
      grd_Listad.Col = 20:       grd_Listad.Text = "OBSERV":    grd_Listad.CellAlignment = flexAlignCenterCenter
        
      For r_int_Contad = 1 To UBound(r_arr_Matriz)
         grd_Listad.Rows = grd_Listad.Rows + 1
         'Numref
         grd_Listad.TextMatrix(r_int_Contad, 0) = r_arr_Matriz(r_int_Contad).CarDocEle_Col1
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr1, r_int_Contad, 0, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Tipo Comprobante
         grd_Listad.TextMatrix(r_int_Contad, 1) = r_arr_Matriz(r_int_Contad).CarDocEle_Col2
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr2, r_int_Contad, 1, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Tipo Proceso
         grd_Listad.TextMatrix(r_int_Contad, 2) = r_arr_Matriz(r_int_Contad).CarDocEle_Col3
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr3, r_int_Contad, 2, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Fecha Emisión
         grd_Listad.TextMatrix(r_int_Contad, 3) = r_arr_Matriz(r_int_Contad).CarDocEle_Col4
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr4, r_int_Contad, 3, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Moneda
         grd_Listad.TextMatrix(r_int_Contad, 4) = r_arr_Matriz(r_int_Contad).CarDocEle_Col5
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr5, r_int_Contad, 4, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Tipo Cambio
         grd_Listad.TextMatrix(r_int_Contad, 5) = r_arr_Matriz(r_int_Contad).CarDocEle_Col6
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr6, r_int_Contad, 5, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'TipDoc
         grd_Listad.TextMatrix(r_int_Contad, 6) = r_arr_Matriz(r_int_Contad).CarDocEle_Col7
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr7, r_int_Contad, 6, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'NumDoc
         grd_Listad.TextMatrix(r_int_Contad, 7) = r_arr_Matriz(r_int_Contad).CarDocEle_Col8
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr8, r_int_Contad, 7, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Receptor - Razón Social
         grd_Listad.TextMatrix(r_int_Contad, 8) = r_arr_Matriz(r_int_Contad).CarDocEle_Col20
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr20, r_int_Contad, 8, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Dirección
         grd_Listad.TextMatrix(r_int_Contad, 9) = r_arr_Matriz(r_int_Contad).CarDocEle_Col9
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr9, r_int_Contad, 9, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Distrito
         grd_Listad.TextMatrix(r_int_Contad, 10) = r_arr_Matriz(r_int_Contad).CarDocEle_Col10
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr10, r_int_Contad, 10, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Provincia
         grd_Listad.TextMatrix(r_int_Contad, 11) = r_arr_Matriz(r_int_Contad).CarDocEle_Col11
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr11, r_int_Contad, 11, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Departamento
         grd_Listad.TextMatrix(r_int_Contad, 12) = r_arr_Matriz(r_int_Contad).CarDocEle_Col12
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr12, r_int_Contad, 12, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Correo
         grd_Listad.TextMatrix(r_int_Contad, 13) = r_arr_Matriz(r_int_Contad).CarDocEle_Col13
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr13, r_int_Contad, 13, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Cantidad
         grd_Listad.TextMatrix(r_int_Contad, 14) = r_arr_Matriz(r_int_Contad).CarDocEle_Col14
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr14, r_int_Contad, 14, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Código
         grd_Listad.TextMatrix(r_int_Contad, 15) = r_arr_Matriz(r_int_Contad).CarDocEle_Col15
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr15, r_int_Contad, 15, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Unidad Medida
         grd_Listad.TextMatrix(r_int_Contad, 16) = r_arr_Matriz(r_int_Contad).CarDocEle_Col16
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr16, r_int_Contad, 16, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Glosa
         grd_Listad.TextMatrix(r_int_Contad, 17) = r_arr_Matriz(r_int_Contad).CarDocEle_Col17
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr17, r_int_Contad, 17, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Importe Unitario
         grd_Listad.TextMatrix(r_int_Contad, 18) = r_arr_Matriz(r_int_Contad).CarDocEle_Col18
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr18, r_int_Contad, 18, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Importe Total Venta
         grd_Listad.TextMatrix(r_int_Contad, 19) = r_arr_Matriz(r_int_Contad).CarDocEle_Col19
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr19, r_int_Contad, 19, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
         
         'Observación
         grd_Listad.TextMatrix(r_int_Contad, 20) = r_arr_Matriz(r_int_Contad).CarDocEle_Col21
         Call fs_Colorear_Celda(r_arr_Matriz(r_int_Contad).CarDocEle_ColErr21, r_int_Contad, 20, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarDocEle_Col1)
      Next r_int_Contad
       
      grd_Listad.Redraw = True
      fs_Cargar_ArcDocEle = True
       
      grd_Listad.Row = 0
      Call gs_RefrescaGrid(grd_Listad)
      Call fs_Activa(True)
   End If
   
Salir:
   r_obj_Excel.Quit
   Set r_obj_Excel = Nothing
End Function

Private Function fs_Validar_Referencia(ByVal p_NumRef As String) As Boolean
   fs_Validar_Referencia = False
   p_NumRef = Replace(p_NumRef, "-", "")
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAECFI_NUMREF AS REFERENCIA "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "   WHERE A.MAECFI_NUMREF = '" & p_NumRef & "'"
     
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   g_rst_Genera.MoveFirst
   If Not g_rst_Genera.EOF Then
      fs_Validar_Referencia = True
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Function fs_Verificar_Cabecera(ByVal r_obj_FicExc As Excel.Application) As Boolean
Dim r_int_NumFil As Integer

   fs_Verificar_Cabecera = False
   r_int_NumFil = 1
   
   If Trim(r_obj_FicExc.Cells(r_int_NumFil, 1).Value) <> "" Then
      If Trim(r_obj_FicExc.Cells(r_int_NumFil, 1).Value) = "NUMREF" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 2).Value) = "TIPCOM" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 3).Value) = "TIPPRO" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 4).Value) = "FECEMI" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 5).Value) = "MONEDA" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 6).Value) = "TIPCAM" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 7).Value) = "TIPDOC" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 8).Value) = "NUMDOC" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 9).Value) = "DIRECCION" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 10).Value) = "DISTRITO" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 11).Value) = "PROVINCIA" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 12).Value) = "DEPARTAMENTO" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 13).Value) = "CORREO" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 14).Value) = "CANTIDAD" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 15).Value) = "CODIGO" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 16).Value) = "UM" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 17).Value) = "GLOSA" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 18).Value) = "VALUNI" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 19).Value) = "VALVTA" Then
         fs_Verificar_Cabecera = True
      ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 20).Value) = "OBSERV" Then
         fs_Verificar_Cabecera = True
      Else
         fs_Verificar_Cabecera = False
      End If
   End If
End Function

Private Function fs_Validar_TipPro(ByVal p_TipPro As Integer) As Boolean
   fs_Validar_TipPro = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT TRIM(A.PARDES_DESCRI) AS DESCRIPCION "
   g_str_Parame = g_str_Parame & "    FROM MNT_PARDES A "
   g_str_Parame = g_str_Parame & "   WHERE PARDES_CODGRP = '539' "
   g_str_Parame = g_str_Parame & "     AND A.PARDES_CODITE = " & p_TipPro & ""
   g_str_Parame = g_str_Parame & "     AND A.PARDES_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   ORDER BY PARDES_CODITE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   g_rst_Genera.MoveFirst
   If Not g_rst_Genera.EOF Then
      fs_Validar_TipPro = True
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function

Private Function fs_Validar_NumDoc(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As Boolean
   fs_Validar_NumDoc = True
   
   If (p_TipDoc = 1) Then 'DNI - 8
      If Len(Trim(p_NumDoc)) <> 8 Then
         MsgBox "El documento de identidad es de 8 digitos.", vbExclamation, modgen_g_str_NomPlt
'         If l_int_TipOpc <> 1 Then
'            Call gs_SetFocus(txt_NumDoc)
'         End If
         fs_Validar_NumDoc = False
      End If
   ElseIf (p_TipDoc = 6) Then 'RUC - 11
      If Not gf_Valida_RUC(Trim(p_NumDoc), Mid(Trim(p_NumDoc), Len(Trim(p_NumDoc)), 1)) Then
         MsgBox "El Número de RUC no es válido " & p_NumDoc, vbExclamation, modgen_g_str_NomPlt
'         If l_int_TipOpc <> 1 Then
'            Call gs_SetFocus(txt_NumDoc)
'         End If
         fs_Validar_NumDoc = False
      End If
   Else 'OTROS
      If Len(Trim(p_NumDoc)) = 0 Then
         MsgBox "Debe ingresar un numero de documento.", vbExclamation, modgen_g_str_NomPlt
'         If l_int_TipOpc <> 1 Then
'            Call gs_SetFocus(txt_NumDoc)
'         End If
         fs_Validar_NumDoc = False
      End If
   End If
End Function

Private Function fs_Buscar_Receptor(ByVal p_TipPro As Integer, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByRef p_RazSoc As String, ByRef p_Direcc As String, ByRef p_Distri As String, ByRef p_Provin As String, ByRef p_Depart As String, ByRef p_Correo As String)
Dim r_int_TipPro     As Integer

   p_RazSoc = ""
   p_Direcc = ""
   p_Distri = ""
   p_Provin = ""
   p_Correo = ""
   p_Depart = ""
   g_str_Parame = ""
   
   If p_TipPro = 3 Or p_TipPro = 4 Or p_TipPro = 6 Then
      
      g_str_Parame = g_str_Parame & " SELECT MAEETE_TIPDOC, MAEETE_NUMDOC, MAEPRV_RAZSOC AS RECEPTOR, TRIM(MAEETE_DIRREP) AS DIRECCION, TRIM(C.PARDES_DESCRI) AS DEPARTAMENTO, "
      g_str_Parame = g_str_Parame & "        TRIM(D.PARDES_DESCRI) AS PROVINCIA, TRIM(E.PARDES_DESCRI) AS DISTRITO, TRIM(MAEPRV_CORREO) AS CORREO "
      g_str_Parame = g_str_Parame & "   FROM TPR_MAEETE A  "
      g_str_Parame = g_str_Parame & "        INNER JOIN CNTBL_MAEPRV B ON B.MAEPRV_TIPDOC = MAEETE_TIPDOC AND B.MAEPRV_NUMDOC = MAEETE_NUMDOC "
      g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES C ON C.PARDES_CODGRP = 101 AND C.PARDES_CODITE = SUBSTR(A.MAEETE_UBIGEO,1,2)||'0000' "
      g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 101 AND D.PARDES_CODITE = SUBSTR(A.MAEETE_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "         LEFT JOIN MNT_PARDES E ON E.PARDES_CODGRP = 101 AND E.PARDES_CODITE = A.MAEETE_UBIGEO "
      g_str_Parame = g_str_Parame & "  WHERE MAEETE_SITUAC = 1  "
      If p_TipDoc > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEETE_TIPDOC = " & p_TipDoc & "  "
      End If
      If Len(Trim(p_NumDoc)) > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEETE_NUMDOC = '" & Trim(p_NumDoc) & "' "
      End If
      
   Else
      g_str_Parame = g_str_Parame & " SELECT MAEPRV_TIPDOC, MAEPRV_NUMDOC, MAEPRV_RAZSOC AS RECEPTOR, TRIM(MAEPRV_DOMFIS) AS DIRECCION, TRIM(B.PARDES_DESCRI) AS DEPARTAMENTO, "
      g_str_Parame = g_str_Parame & "        TRIM(C.PARDES_DESCRI) AS PROVINCIA, TRIM(D.PARDES_DESCRI) AS DISTRITO, TRIM(MAEPRV_CORREO) AS CORREO "
      g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A  "
      g_str_Parame = g_str_Parame & "        LEFT JOIN MNT_PARDES B ON B.PARDES_CODGRP = 101 AND B.PARDES_CODITE = SUBSTR(A.MAEPRV_UBIGEO,1,2)||'0000' "
      g_str_Parame = g_str_Parame & "        LEFT JOIN MNT_PARDES C ON C.PARDES_CODGRP = 101 AND C.PARDES_CODITE = SUBSTR(A.MAEPRV_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "        LEFT JOIN MNT_PARDES D ON D.PARDES_CODGRP = 101 AND D.PARDES_CODITE = A.MAEPRV_UBIGEO "
      g_str_Parame = g_str_Parame & "  WHERE MAEPRV_SITUAC = 1  "
      If p_TipDoc > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEPRV_TIPDOC = " & p_TipDoc & "  "
      End If
      If Len(Trim(p_NumDoc)) > 0 Then
         g_str_Parame = g_str_Parame & "   AND MAEPRV_NUMDOC = '" & Trim(p_NumDoc) & "' "
      End If
      
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   g_rst_Princi.MoveFirst
   If Not g_rst_Princi.EOF Then
      If Not IsNull(g_rst_Princi!RECEPTOR) Then
         p_RazSoc = Trim(g_rst_Princi!RECEPTOR)
      End If
      If Not IsNull(g_rst_Princi!Direccion) Then
         p_Direcc = Trim(g_rst_Princi!Direccion)
      End If
      If Not IsNull(g_rst_Princi!DEPARTAMENTO) Then
         p_Depart = Trim(g_rst_Princi!DEPARTAMENTO)
      End If
      If Not IsNull(g_rst_Princi!PROVINCIA) Then
         p_Provin = Trim(g_rst_Princi!PROVINCIA)
      End If
      If Not IsNull(g_rst_Princi!DISTRITO) Then
         p_Distri = Trim(g_rst_Princi!DISTRITO)
      End If
      If Not IsNull(g_rst_Princi!CORREO) Then
         p_Correo = Trim(g_rst_Princi!CORREO)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

Private Sub fs_Activa(ByVal p_Estado As Boolean)
'   txt_NomArc.Enabled = p_Estado
'   cmd_BuscaArc.Enabled = p_Estado
   cmd_Grabar.Enabled = p_Estado
   cmd_Cancel.Enabled = p_Estado
'   cmd_Import.Enabled = p_Estado
End Sub

Private Sub fs_Colorear_Celda(ByVal r_str_DesErr As String, ByVal r_int_NumFil As Integer, ByVal r_int_NumCol As Integer, ByRef r_arr_Mat() As g_tpo_CarDocEle, ByVal r_int_NumRefer As Integer)
Dim r_int_ConAux   As Integer

   If Not IsNull(r_str_DesErr) And r_str_DesErr <> "" Then
      If r_int_NumFil = 0 Then
          For r_int_ConAux = 1 To UBound(r_arr_Mat)
              If r_arr_Mat(r_int_ConAux).CarDocEle_Col1 = r_int_NumRefer Then
                  grd_Listad.Row = r_int_ConAux
                  grd_Listad.Col = r_int_NumCol
                  grd_Listad.CellBackColor = &H8080FF
                  l_bol_FlgErr = True
              End If
          Next r_int_ConAux
      Else
          If r_arr_Mat(r_int_NumFil).CarDocEle_Col1 = r_int_NumRefer Then
              grd_Listad.Row = r_int_NumFil
              grd_Listad.Col = r_int_NumCol
              grd_Listad.CellBackColor = &H8080FF
              l_bol_FlgErr = True
          End If
      End If
    End If
End Sub

Private Function fs_Agrega_Grilla_DocEle() As Boolean
Dim r_int_Contad        As Integer
Dim r_int_ConAux        As Integer
Dim r_int_NumRef        As Integer
Dim r_int_RefAnt        As Integer

'Cabecera
'Dim r_str_PerAnn        As String
'Dim r_str_PerMes        As String
Dim r_str_TipCom        As String
Dim r_str_NumCom        As String
Dim r_int_TipPro        As Integer
Dim r_str_Fecemi        As String
Dim r_str_moneda        As String
Dim r_dbl_TipCam        As Double
Dim r_int_TipDoc        As Integer
Dim r_str_NumDoc        As String
Dim r_str_Direcc        As String
Dim r_str_Distri        As String
Dim r_str_Provin        As String
Dim r_str_Depart        As String
Dim r_str_Correo        As String

'Detalle
Dim r_int_Cantid        As Integer
Dim r_str_Codigo        As String
Dim r_str_Unidad        As String
Dim r_str_GloDet        As String
Dim r_dbl_MtoUni        As Double
Dim r_dbl_MtoVta        As Double
Dim r_int_NumItm        As Integer

Dim r_int_SerFac        As Integer
Dim r_lng_NumFac        As Long
Dim r_str_NumSer        As String
Dim r_str_NumFac        As String
Dim r_str_RazSoc        As String
Dim r_str_Observ        As String

Dim r_arr_Matriz()      As g_tpo_CarDocEle

   Screen.MousePointer = 11
   moddat_g_int_FlgGrb = 1
   l_int_RegTot = 0
   l_int_RegPro = 0
   l_int_RegSPr = 0
   l_int_RegErr = 0
   l_bol_FlgErr = False
    
   fs_Agrega_Grilla_DocEle = False
'   grd_Listad.Redraw = False
   
'   ReDim l_arr_LogPro(0)
'   ReDim l_arr_LogPro(1)
   ReDim r_arr_Matriz(0)
   
'   l_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
'   l_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
'   l_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
'   l_arr_LogPro(1).LogPro_NumErr = 0
        
   'Total de registros del archivo
   For r_int_Contad = 1 To grd_Listad.Rows - 1
'      r_int_NumRef = Trim(grd_Listad.TextMatrix(r_int_Contad, 0))
      If r_int_RefAnt <> r_int_NumRef Then
         l_int_RegTot = l_int_RegTot + 1
'         r_int_RefAnt = r_int_NumRef
      End If
   Next r_int_Contad
   
   r_int_RefAnt = 0
   r_int_NumRef = 0
   l_lng_Codigo = 0
   
   For r_int_Contad = 1 To grd_Listad.Rows - 1
   
      r_int_NumRef = Trim(grd_Listad.TextMatrix(r_int_Contad, 0))
      r_str_TipCom = Trim(grd_Listad.TextMatrix(r_int_Contad, 1))
      If r_str_TipCom = "F" Then
         r_str_NumCom = "01"
      ElseIf r_str_TipCom = "B" Then
         r_str_NumCom = "03"
      End If
      r_int_TipPro = Mid(Trim(grd_Listad.TextMatrix(r_int_Contad, 2)), 1, 1)
      r_str_Fecemi = Trim(grd_Listad.TextMatrix(r_int_Contad, 3))
      r_str_moneda = fs_Obtener_Moneda(Trim(grd_Listad.TextMatrix(r_int_Contad, 4)))
      r_dbl_TipCam = Trim(grd_Listad.TextMatrix(r_int_Contad, 5))
      r_int_TipDoc = Mid(Trim(grd_Listad.TextMatrix(r_int_Contad, 6)), 1, 1)
      r_str_NumDoc = Trim(grd_Listad.TextMatrix(r_int_Contad, 7))
      r_str_RazSoc = UCase(Trim(grd_Listad.TextMatrix(r_int_Contad, 8)))
      r_str_Direcc = UCase(Trim(grd_Listad.TextMatrix(r_int_Contad, 9)))
      r_str_Distri = UCase(Trim(grd_Listad.TextMatrix(r_int_Contad, 10)))
      r_str_Provin = UCase(Trim(grd_Listad.TextMatrix(r_int_Contad, 11)))
      r_str_Depart = UCase(Trim(grd_Listad.TextMatrix(r_int_Contad, 12)))
      r_str_Correo = Trim(grd_Listad.TextMatrix(r_int_Contad, 13))
      r_int_Cantid = Trim(grd_Listad.TextMatrix(r_int_Contad, 14))
      r_str_Codigo = Trim(grd_Listad.TextMatrix(r_int_Contad, 15))
      r_str_Unidad = Trim(grd_Listad.TextMatrix(r_int_Contad, 16))
      r_str_GloDet = Trim(grd_Listad.TextMatrix(r_int_Contad, 17))
      r_dbl_MtoUni = Trim(grd_Listad.TextMatrix(r_int_Contad, 18))
      r_dbl_MtoVta = Trim(grd_Listad.TextMatrix(r_int_Contad, 19))
      r_str_Observ = UCase(Trim(grd_Listad.TextMatrix(r_int_Contad, 20)))
         
      ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col1 = r_int_NumRef
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col2 = r_str_NumCom
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col3 = r_int_TipPro
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col4 = r_str_Fecemi
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col5 = r_str_moneda
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col6 = r_dbl_TipCam
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col7 = r_int_TipDoc
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col8 = r_str_NumDoc
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col9 = r_str_Direcc
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col10 = r_str_Distri
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col11 = r_str_Provin
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col12 = r_str_Depart
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col13 = r_str_Correo
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col14 = r_int_Cantid
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col15 = r_str_Codigo
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col16 = r_str_Unidad
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col17 = r_str_GloDet
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col18 = r_dbl_MtoUni
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col19 = r_dbl_MtoVta
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col20 = r_str_RazSoc
      r_arr_Matriz(UBound(r_arr_Matriz)).CarDocEle_Col21 = r_str_Observ
      
   Next r_int_Contad
   
   If UBound(r_arr_Matriz) > 0 Then
   Screen.MousePointer = 11
      If fs_Agregar_DocEleTmp(r_arr_Matriz()) = True Then
         r_int_NumItm = r_int_NumItm + 1
         r_int_RefAnt = r_int_NumRef
      Else
         l_bol_FlgErr = True
         l_str_MsjRef = r_int_NumRef
      End If
   End If

   l_int_RegPro = r_int_ConAux
'   grd_Listad.Redraw = True
   Screen.MousePointer = 0
   
   Call gs_RefrescaGrid(grd_Listad)
   Call fs_Limpiar
'   Call gs_SetFocus(cmd_BuscaArc)
   
   fs_Agrega_Grilla_DocEle = True
   l_int_RegErr = IIf(l_bol_FlgErr = True, 1, 0)
   l_int_RegSPr = l_int_RegTot - l_int_RegPro - l_int_RegErr
   l_str_MsjRef = IIf(l_bol_FlgErr = True, "NumRef: " & l_str_MsjRef, "")
   
End Function
Private Function fs_Obtener_Moneda(ByVal p_Moneda As String) As String
Dim r_str_Parame  As String

   fs_Obtener_Moneda = ""
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT CATSUN_CODIGO "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_CATSUN "
   r_str_Parame = r_str_Parame & "  WHERE CATSUN_DESCRI = '" & Trim(p_Moneda) & "' "
   r_str_Parame = r_str_Parame & "    AND CATSUN_NROCAT = 2 "
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
       g_rst_Genera.Close
       Set g_rst_Genera = Nothing
       Exit Function
   Else
       fs_Obtener_Moneda = g_rst_Genera!CATSUN_CODIGO
   End If
End Function
Private Sub fs_Inicia()
 
   'Datos de documentos electrónicos
   grd_Listad.ColWidth(0) = 900
   grd_Listad.ColWidth(1) = 900
   grd_Listad.ColWidth(2) = 3000
   grd_Listad.ColWidth(3) = 900
   grd_Listad.ColWidth(4) = 900
   grd_Listad.ColWidth(5) = 900
   grd_Listad.ColWidth(6) = 3000
   grd_Listad.ColWidth(7) = 1250
   grd_Listad.ColWidth(8) = 3000
   grd_Listad.ColWidth(9) = 3000
   grd_Listad.ColWidth(10) = 2000
   grd_Listad.ColWidth(11) = 2000 '1550
   grd_Listad.ColWidth(12) = 2000
   grd_Listad.ColWidth(13) = 2500
   grd_Listad.ColWidth(14) = 1250
   grd_Listad.ColWidth(15) = 1250
   grd_Listad.ColWidth(16) = 1250
   grd_Listad.ColWidth(17) = 3000
   grd_Listad.ColWidth(18) = 1250
   grd_Listad.ColWidth(19) = 1250
   grd_Listad.ColWidth(20) = 3000
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignLeftCenter
   grd_Listad.ColAlignment(8) = flexAlignLeftCenter
   grd_Listad.ColAlignment(9) = flexAlignLeftCenter
   grd_Listad.ColAlignment(10) = flexAlignLeftCenter
   grd_Listad.ColAlignment(11) = flexAlignLeftCenter
   grd_Listad.ColAlignment(12) = flexAlignLeftCenter
   grd_Listad.ColAlignment(13) = flexAlignLeftCenter
   grd_Listad.ColAlignment(15) = flexAlignCenterCenter
   grd_Listad.ColAlignment(16) = flexAlignCenterCenter
   grd_Listad.ColAlignment(17) = flexAlignLeftCenter
   grd_Listad.ColAlignment(18) = flexAlignRightCenter
   grd_Listad.ColAlignment(19) = flexAlignRightCenter
   grd_Listad.ColAlignment(20) = flexAlignLeftCenter
   grd_Listad.Rows = 0
End Sub
Private Function fs_Agregar_DocEleTmp(ByRef p_Array() As g_tpo_CarDocEle) As Boolean
Dim r_lng_Contad     As Long
Dim r_int_SerFac     As Integer
Dim r_str_NumSer     As String
Dim r_str_NumFac     As String
Dim r_lng_NumFac     As Long
Dim r_str_TipCom     As String

      fs_Agregar_DocEleTmp = False
            
      For r_lng_Contad = 1 To UBound(p_Array)
         
         Call fs_Obtener_Codigo(l_lng_Codigo)
         r_str_TipCom = IIf(Format(p_Array(r_lng_Contad).CarDocEle_Col2, "00") = 1, "F", "B")
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELETMP (      "
         g_str_Parame = g_str_Parame & " DOCELETMP_CODIGO                 , "
         g_str_Parame = g_str_Parame & " DOCELETMP_FECPRO                 , "
         g_str_Parame = g_str_Parame & " DOCELETMP_FECAPR                 , "
         g_str_Parame = g_str_Parame & " DOCELETMP_IDE_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_IDE_FECEMI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_IDE_HOREMI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_IDE_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_IDE_TIPMON             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_IDE_NUMORC             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_IDE_FECVCT             , "
         
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_NOMCOM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_UBIGEO             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_URBANI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_TELEMI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_EMI_COREMI             , "
         
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_TELREC             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REC_CORREC             , "
         
         g_str_Parame = g_str_Parame & " DOCELETMP_DRF_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DRF_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DRF_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DRF_CODMOT             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DRF_DESMOT             , "
         
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTVTA_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTVTA_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTVTA_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTVTA_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_OPEEXP      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTVTA_OPEEXP      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_REGPER      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_BASIMP_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_MTOPER             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_MTOTOT_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIMP             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_MTOIMP             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_OTRCAR             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_CODIGO_TOTDSC      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTDSC             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_IMPTOT_DOCUME      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_DSCGLO             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_INFPPG             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TOTANT             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_TIPOPE             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_CAB_LEYEND             , "
         
         g_str_Parame = g_str_Parame & " DOCELETMP_ADI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_ADI_TITADI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_ADI_VALADI             , "
         
         'DETALLE
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_NUMITE             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODPRD             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_DESPRD             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CANTID             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_UNIDAD             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_VALUNI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_PUNVTA             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIMP             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_MTOIMP             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_TIPAFE             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_VALVTA             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_VALREF             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_DSTITE             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_NUMPLA             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODSUN             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODCON             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_NROCON             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIGO_FECOTO      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_FECOTO             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIGO_TIPPRE      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_TIPPRE             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIGO_PARREG      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_PARREG             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIGO_PRIVIV      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_PRIVIV             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODIGO_DIRCOM      , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODUBI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_UBIGEO             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODURB             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_URBANI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODDPT             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODPRV             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_CODDIS             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_DET_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELETMP_TIPCAM                 , "
         g_str_Parame = g_str_Parame & " DOCELETMP_SITUAC                 , "
         g_str_Parame = g_str_Parame & " DOCELETMP_TIPPRO                 , "
         g_str_Parame = g_str_Parame & " DOCELETMP_REFER                  , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                        , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                        , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                        , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                        , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                        , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         
         g_str_Parame = g_str_Parame & " VALUES ( "
         g_str_Parame = g_str_Parame & "" & l_lng_Codigo & "                                                   , "         'DOCELETMP_CODIGO
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "                      , "         'DOCELETMP_FECPRO
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_FECAPR
         g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'                                                 , "         'DOCELETMP_IDE_SERNUM ----  r_str_NumSer & "-" & r_str_NumFac &
         g_str_Parame = g_str_Parame & "'" & p_Array(r_lng_Contad).CarDocEle_Col4 & "'                         , "         'DOCELETMP_IDE_FECEMI
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_IDE_HOREMI
         g_str_Parame = g_str_Parame & "'" & Format(p_Array(r_lng_Contad).CarDocEle_Col2, "00") & "'           , "         'DOCELETMP_IDE_TIPDOC
         g_str_Parame = g_str_Parame & "'" & p_Array(r_lng_Contad).CarDocEle_Col5 & "'                         , "         'DOCELETMP_IDE_TIPMON
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_IDE_NUMORC
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_IDE_FECVCT
         
         g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "                                             'DOCELETMP_EMI_SERNUM  ---- r_str_NumSer & "-" & r_str_NumFac &
         g_str_Parame = g_str_Parame & "'6'                                                                    , "         'DOCELETMP_EMI_TIPDOC
         g_str_Parame = g_str_Parame & "'20511904162'                                                          , "         'DOCELETMP_EMI_NUMDOC
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_NOMCOM
         g_str_Parame = g_str_Parame & "'EDPYME MICASITA SA'                                                   , "         'DOCELETMP_EMI_DENOMI
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_UBIGEO
         g_str_Parame = g_str_Parame & "'AV RIVERA NAVARRETE 645'                                              , "         'DOCELETMP_EMI_DIRCOM
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_URBANI
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_PROVIN
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_DEPART
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_DISTRI
         g_str_Parame = g_str_Parame & "'PE'                                                                   , "         'DOCELETMP_EMI_CODPAI
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_TELEMI
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_EMI_COREMI
         
         g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "                                             'DOCELETMP_REC_SERNUM ----   r_str_NumSer & "-" & r_str_NumFac &
         g_str_Parame = g_str_Parame & "'" & p_Array(r_lng_Contad).CarDocEle_Col7 & "'                         , "         'DOCELETMP_REC_TIPDOC
         g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col8) & "'                   , "         'DOCELETMP_REC_NUMDOC
         g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col20) & "'                  , "         'DOCELETMP_REC_DENOMI
         If Len(p_Array(r_lng_Contad).CarDocEle_Col9) = 0 Then
            g_str_Parame = g_str_Parame & " NULL, "                                                                        'DOCELETMP_REC_DIRCOM
         Else
            g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col9) & "'                , "
         End If
         If Len(p_Array(r_lng_Contad).CarDocEle_Col10) = 0 Then
            g_str_Parame = g_str_Parame & " NULL, "                                                                        'DOCELETMP_REC_DISTRI
         Else
            g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col10) & "'               , "
         End If
         If Len(p_Array(r_lng_Contad).CarDocEle_Col11) = 0 Then
            g_str_Parame = g_str_Parame & " NULL, "                                                                        'DOCELETMP_REC_PROVIN
         Else
            g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col11) & "'               , "
         End If
         If Len(p_Array(r_lng_Contad).CarDocEle_Col12) = 0 Then
            g_str_Parame = g_str_Parame & " NULL, "                                                                        'DOCELETMP_REC_DEPART
         Else
            g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col12) & "'               , "
         End If
         g_str_Parame = g_str_Parame & "'PE'                                                                   , "         'DOCELETMP_REC_CODPAI
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_REC_TELREC
         If Len(p_Array(r_lng_Contad).CarDocEle_Col13) = 0 Then
            g_str_Parame = g_str_Parame & " NULL, "                                                                        'DOCELETMP_REC_CORREC
         Else
            g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col13) & "'               , "
         End If
         g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "                                             'DOCELETMP_DRF_SERNUM  ---- r_str_NumSer & "-" & r_str_NumFac &
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DRF_TIPDOC
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DRF_NUMDOC
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DRF_CODMOT
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DRF_DESMOT
         
         g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "          'DOCELETMP_CAB_SERNUM  ---- & r_str_NumSer & "-" & r_str_NumFac &
         g_str_Parame = g_str_Parame & "'1001'                                                                 , "         'DOCELETMP_CAB_CODIGO_OPEGRV
         g_str_Parame = g_str_Parame & " 0 , "                                                                             'DOCELETMP_CAB_TOTVTA_OPEGRV
         g_str_Parame = g_str_Parame & "'1002'                                                                 , "         'DOCELETMP_CAB_CODIGO_OPEINA
         g_str_Parame = g_str_Parame & "" & p_Array(r_lng_Contad).CarDocEle_Col18 & "                          , "         'DOCELETMP_CAB_TOTVTA_OPEINA
         g_str_Parame = g_str_Parame & "'1003'                                                                 , "         'DOCELETMP_CAB_CODIGO_OPEEXO
         g_str_Parame = g_str_Parame & " 0 , "                                                                             'DOCELETMP_CAB_TOTVTA_OPEEXO
         g_str_Parame = g_str_Parame & "'1004'                                                                 , "         'DOCELETMP_CAB_CODIGO_OPEGRA
         g_str_Parame = g_str_Parame & " 0 , "                                                                             'DOCELETMP_CAB_TOTVTA_OPEGRA
         g_str_Parame = g_str_Parame & "'1000'                                                                 , "         'DOCELETMP_CAB_CODIGO_OPEEXP
         g_str_Parame = g_str_Parame & " 0 , "                                                                             'DOCELETMP_CAB_TOTVTA_OPEEXP
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_CODIGO_PERCEP
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_CODIGO_REGPER
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_BASIMP_PERCEP
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_MTOPER
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_MTOTOT_PERCEP
         g_str_Parame = g_str_Parame & "'1000'                                                                 , "         'DOCELETMP_CAB_CODIMP
         g_str_Parame = g_str_Parame & " 0 , "                                                                             'DOCELETMP_CAB_MTOIMP
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_OTRCAR
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_CODIGO_TOTDSC
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_TOTDSC
         g_str_Parame = g_str_Parame & "" & p_Array(r_lng_Contad).CarDocEle_Col19 & "                          , "         'DOCELETMP_CAB_IMPTOT_DOCUME
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_DSCGLO
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_INFPPG
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_CAB_TOTANT
         g_str_Parame = g_str_Parame & "'13'                                                                   , "         'DOCELETMP_CAB_TIPOPE  -- Verificar CATALOGO 17 -- Gasto Deducible Persona Natural
         g_str_Parame = g_str_Parame & "'[1000'                                                                , "         'DOCELETMP_CAB_LEYEND
         
         g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "          'DOCELETMP_ADI_SERNUM ---- r_str_NumSer & "-" & r_str_NumFac &
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_ADI_TITADI
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_ADI_VALADI
         
         'DETALLE
         g_str_Parame = g_str_Parame & "'" & r_str_TipCom & "'             , "         'DOCELETMP_DET_SERNUM ---- r_str_NumSer & "-" & r_str_NumFac &
         g_str_Parame = g_str_Parame & "'001'                                                                  , "         'DOCELETMP_DET_NUMITE
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODPRD
         g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col17) & "'                  , "         'DOCELETMP_DET_DESPRD
         g_str_Parame = g_str_Parame & "" & p_Array(r_lng_Contad).CarDocEle_Col14 & "                          , "         'DOCELETMP_DET_CANTID
         g_str_Parame = g_str_Parame & "'" & CStr(p_Array(r_lng_Contad).CarDocEle_Col16) & "'                  , "         'DOCELETMP_DET_UNIDAD
         g_str_Parame = g_str_Parame & "" & CDbl(p_Array(r_lng_Contad).CarDocEle_Col18) & "                    , "         'DOCELETMP_DET_VALUNI
         g_str_Parame = g_str_Parame & "" & CDbl(p_Array(r_lng_Contad).CarDocEle_Col19) & "                    , "         'DOCELETMP_DET_PUNVTA
         g_str_Parame = g_str_Parame & "'1000'                                                                 , "         'DOCELETMP_DET_CODIMP
         g_str_Parame = g_str_Parame & "0.00                                                                   , "         'DOCELETMP_DET_MTOIMP
         g_str_Parame = g_str_Parame & "'30'                                                                   , "         'DOCELETMP_DET_TIPAFE -   'Inafecto operación onerosa
         g_str_Parame = g_str_Parame & "" & CDbl(p_Array(r_lng_Contad).CarDocEle_Col19) & "                    , "         'DOCELETMP_DET_VALVTA
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_VALREF
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_DSTITE
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_NUMPLA
         
         g_str_Parame = g_str_Parame & "'84121501'                                                             , "         'DOCELETMP_DET_CODSUN
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODCON
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_NROCON
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODIGO_FECOTO
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_FECOTO
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODIGO_TIPPRE
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_TIPPRE
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODIGO_PARREG
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_PARREG
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODIGO_PRIVIV
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_PRIVIV
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODIGO_DIRCOM
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_DIRCOM
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODUBI
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_UBIGEO
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODURB
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_URBANI
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODDPT
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_DEPART
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODPRV
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_PROVIN
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_CODDIS
         g_str_Parame = g_str_Parame & " NULL, "                                                                           'DOCELETMP_DET_DISTRI
         
         g_str_Parame = g_str_Parame & "" & CDbl(p_Array(r_lng_Contad).CarDocEle_Col6) & "                     , "         'DOCELETMP_TIPCAM
         g_str_Parame = g_str_Parame & "" & 2 & "                                                              , "         'DOCELETMP_SITUAC
         g_str_Parame = g_str_Parame & "" & p_Array(r_lng_Contad).CarDocEle_Col3 & "                           , "         'DOCELETMP_TIPPRO
         g_str_Parame = g_str_Parame & "'" & p_Array(r_lng_Contad).CarDocEle_Col21 & "'                        , "         'DOCELETMP_REFER
  
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "'                                          , "         'SEGUSUCRE
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "                , "         'SEGFECCRE
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                                        , "         'SEGHORCRE
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                                           , "         'SEGPLTCRE
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                                          , "         'SEGTERCRE
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                               
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELETMP, Nro Doc:" & CStr(p_Array(r_lng_Contad).CarDocEle_Col8) & ", Nro. Cod: " & l_lng_Codigo & ", procedimiento: fs_Agregar_DocEleTmp")
            Exit Function
         End If
         DoEvents: DoEvents: DoEvents
         
       Set g_rst_Genera = Nothing
   Next
   fs_Agregar_DocEleTmp = True
End Function
Private Function fs_Obtener_Codigo(ByRef p_CodIte As Long)
Dim r_str_Parame           As String
Dim r_rst_Codigo           As ADODB.Recordset
   
   p_CodIte = 0
     
   'Código Máximo de CNTBL_DOCELE
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT MAX(DOCELETMP_CODIGO) AS CODIGO "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_DOCELETMP "
      
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para obtener codigo en CNTBL_DOCELE, procedimiento: fs_ObtenerCodigo")
      Exit Function
   End If
   
   If r_rst_Codigo.BOF And r_rst_Codigo.EOF Then
      r_rst_Codigo.Close
      Set r_rst_Codigo = Nothing
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontro ningún registro en CNTBL_DOCELE, procedimiento: fs_ObtenerCodigo")
      Exit Function
   End If
   
   If Not (r_rst_Codigo.BOF And r_rst_Codigo.EOF) Then
      r_rst_Codigo.MoveFirst
      If IsNull(r_rst_Codigo!CODIGO) Then
         p_CodIte = 0
      Else
         p_CodIte = r_rst_Codigo!CODIGO
      End If
   End If
   
   p_CodIte = p_CodIte + 1
   
   r_rst_Codigo.Close
   Set r_rst_Codigo = Nothing

End Function
Private Sub fs_Escribir_Linea(p_ruta As String, p_texto As String)
   On Error GoTo MyError

   'Escribir en archivo según se ubique
   Set l_fsobj = New FileSystemObject
   Set l_txtStr = l_fsobj.OpenTextFile(p_ruta, ForAppending, False)
   l_txtStr.WriteLine (p_texto)
   l_txtStr.Close
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Escribir_Linea")
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpiar
   Call fs_Activa(False)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub
