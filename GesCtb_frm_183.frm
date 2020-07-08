VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_Ctb_AsiCtb_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14115
   Icon            =   "GesCtb_frm_183.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   14115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel13 
      Height          =   8355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14115
      _Version        =   65536
      _ExtentX        =   24897
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
         TabIndex        =   1
         Top             =   60
         Width           =   13995
         _Version        =   65536
         _ExtentX        =   24686
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
            TabIndex        =   2
            Top             =   30
            Width           =   4755
            _Version        =   65536
            _ExtentX        =   8387
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Carga de Asientos Contables"
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
            Picture         =   "GesCtb_frm_183.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   60
         TabIndex        =   3
         Top             =   720
         Width           =   13995
         _Version        =   65536
         _ExtentX        =   24686
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
            Left            =   13380
            Picture         =   "GesCtb_frm_183.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   60
            Picture         =   "GesCtb_frm_183.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   60
         TabIndex        =   6
         Top             =   1410
         Width           =   13995
         _Version        =   65536
         _ExtentX        =   24686
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
         Begin VB.CommandButton cmd_Import 
            Height          =   585
            Left            =   13395
            Picture         =   "GesCtb_frm_183.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Importar archivo"
            Top             =   30
            Width           =   585
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
            Left            =   12060
            TabIndex        =   8
            ToolTipText     =   "Seleccionar archivo"
            Top             =   180
            Width           =   315
         End
         Begin VB.TextBox txt_NomArc 
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "txt_NomArc"
            Top             =   180
            Width           =   10455
         End
         Begin VB.Label Label4 
            Caption         =   "Archivo a cargar:"
            Height          =   255
            Left            =   180
            TabIndex        =   10
            Top             =   210
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   6165
         Left            =   60
         TabIndex        =   11
         Top             =   2100
         Width           =   13995
         _Version        =   65536
         _ExtentX        =   24686
         _ExtentY        =   10874
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
            Height          =   6015
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   10610
            _Version        =   393216
            Rows            =   30
            Cols            =   15
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
            Appearance      =   0
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_AsiCtb_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type g_tpo_CarAsiCtb
   CarAsiCtb_Col1       As String
   CarAsiCtb_ColErr1    As String
   CarAsiCtb_Col2       As String
   CarAsiCtb_ColErr2    As String
   CarAsiCtb_Col3       As String
   CarAsiCtb_ColErr3    As String
   CarAsiCtb_Col4       As String
   CarAsiCtb_ColErr4    As String
   CarAsiCtb_Col5       As String
   CarAsiCtb_ColErr5    As String
   CarAsiCtb_Col6       As String
   CarAsiCtb_ColErr6    As String
   CarAsiCtb_Col7       As String
   CarAsiCtb_ColErr7    As String
   CarAsiCtb_Col8       As String
   CarAsiCtb_ColErr8    As String
   CarAsiCtb_Col9       As String
   CarAsiCtb_ColErr9    As String
   CarAsiCtb_Col10      As String
   CarAsiCtb_ColErr10   As String
   CarAsiCtb_Col11      As String
   CarAsiCtb_ColErr11   As String
   CarAsiCtb_Col12      As String
   CarAsiCtb_ColErr12   As String
   CarAsiCtb_Col13      As String
   CarAsiCtb_ColErr13   As String
   CarAsiCtb_Col14      As String
   CarAsiCtb_ColErr14   As String
   CarAsiCtb_Col15      As String
   CarAsiCtb_ColErr15   As String
End Type

Private Type g_tpo_VerAsiCtb
   VerAsiCtb_Col1       As Integer
End Type

Dim l_arr_LogPro()      As modprc_g_tpo_LogPro
Dim l_bol_FlgErr        As Boolean
Dim l_int_NumReg        As Integer
Dim l_str_MsjAsi        As String
Dim l_str_MsjRef        As String
Dim l_int_RegTot        As Integer
Dim l_int_RegPro        As Integer
Dim l_int_RegErr        As Integer
Dim l_int_RegSPr        As Integer

Private Sub cmd_BuscaArc_Click()
   dlg_Guarda.Filter = "Archivos Excel |*.xlsx;*.xls"
   dlg_Guarda.ShowOpen
   txt_NomArc.Text = UCase(dlg_Guarda.FileName)
   Exit Sub
End Sub

Private Sub cmd_Grabar_Click()
   If l_bol_FlgErr = False Then
      If MsgBox("¿Desea cargar a la base de datos la información de los asientos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      'proceso de grabación
      If fs_Actualiza_AsientoContable() Then
         Screen.MousePointer = 0
         MsgBox "Proceso realizado satisfactoriamente ", vbInformation, modgen_g_str_NomPlt
      Else
         MsgBox "Registros Encontrados" & vbTab & ":" & "   " & l_int_RegTot & vbNewLine & _
                "Registros Procesados" & vbTab & ":" & "   " & l_int_RegPro & vbNewLine & _
                "Registros Errados" & vbTab & vbTab & ":" & "   " & l_int_RegErr & vbNewLine & _
                "Registros Sin Procesar" & vbTab & ":" & "   " & l_int_RegSPr, vbInformation, modgen_g_str_NomPlt
      End If
   Else
       MsgBox "No se puede cargar la información a la base de datos, aún existen errores.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(grd_Listad)
       Exit Sub
   End If
End Sub

Private Sub cmd_Import_Click()
   'validaciones
   If Len(Trim(txt_NomArc.Text)) = 0 Then
      MsgBox "Debe ingresar la ubicación y nombre del archivo a importar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If
   
   MsgBox "Antes de realizar el proceso de carga verifique los siguiente: " & vbCrLf & " - El archivo excel debe tener el formato del 2007. " & vbCrLf & " - La Columna E y L del archivo con formato 'dd/mm/aaaa'", vbInformation, modgen_g_str_NomPlt
   If MsgBox("¿Desea realizar la carga del archivo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   If fs_Carga_ArchivoAsiento Then
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

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_IniciaGrid
   Call fs_Limpiar
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(txt_NomArc)
   Screen.MousePointer = 0
End Sub

Private Sub fs_IniciaGrid()
   'Datos del Asiento
   grd_Listad.ColWidth(0) = 705
   grd_Listad.ColWidth(1) = 765
   grd_Listad.ColWidth(2) = 765
   grd_Listad.ColWidth(3) = 855
   grd_Listad.ColWidth(4) = 1215
   grd_Listad.ColWidth(5) = 765
   grd_Listad.ColWidth(6) = 855
   grd_Listad.ColWidth(7) = 855
   grd_Listad.ColWidth(8) = 5415
   grd_Listad.ColWidth(9) = 1215
   grd_Listad.ColWidth(10) = 5415
   grd_Listad.ColWidth(11) = 1215
   grd_Listad.ColWidth(12) = 855
   grd_Listad.ColWidth(13) = 1215
   grd_Listad.ColWidth(14) = 1215
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(8) = flexAlignLeftCenter
   grd_Listad.ColAlignment(9) = flexAlignCenterCenter
   grd_Listad.ColAlignment(10) = flexAlignLeftCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
   grd_Listad.ColAlignment(12) = flexAlignCenterCenter
   grd_Listad.Rows = 0
End Sub

Private Sub fs_Limpiar()
   Call gs_LimpiaGrid(grd_Listad)
   txt_NomArc.Text = ""
   Call gs_SetFocus(txt_NomArc)
   cmd_Grabar.Enabled = False
End Sub

Private Function fs_Carga_ArchivoAsiento() As Boolean
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_int_Contad        As Integer
Dim r_int_TotReg        As Integer
Dim r_int_NumRef        As Integer
Dim r_int_RefAnt        As Integer
Dim r_int_ConAux        As Integer

'Cabecera
Dim r_str_PerAnn        As String
Dim r_str_PerMes        As String
Dim r_str_AnnAnt        As String
Dim r_str_MesAnt        As String
Dim r_int_NroLib        As Integer
Dim r_dat_FecCom        As Date
Dim r_str_TipOpe        As String
Dim r_str_TipMon        As String
Dim r_dbl_TipCam        As Double
Dim r_str_GloCab        As String

'Detalle
Dim r_str_CtaCtb        As String
Dim r_str_GloDet        As String
Dim r_dat_FecCtb        As Date
Dim r_str_DebHab        As String
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_arr_Matriz()      As g_tpo_CarAsiCtb
Dim r_arr_MatVal()      As g_tpo_VerAsiCtb
Dim r_dbl_DebSol        As Double
Dim r_dbl_HabSol        As Double
Dim r_dbl_DebDol        As Double
Dim r_dbl_HabDol        As Double
Dim r_bol_FlgExi        As Boolean

   l_bol_FlgErr = False
   fs_Carga_ArchivoAsiento = False
   Call gs_LimpiaGrid(grd_Listad)
   ReDim r_arr_Matriz(0)
   ReDim r_arr_MatVal(0)
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=txt_NomArc.Text
   
   'Valida y Carga Cronograma No Concesional FMV
   r_int_FilExc = 0
   
   If fs_VerificarCabecera(r_obj_Excel) = True Then
      
      Do While Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value) <> ""
         
         ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
           
         'Número de referencia
         If IsNumeric(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)) Then
            r_int_NumRef = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col1 = r_int_NumRef
         Else
            r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col1 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)
            r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr1 = "Número de Referencia Inválido."
         End If
         
         'Valida Año y Mes actual
         r_str_PerAnn = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value)
         r_str_PerMes = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
         
         If r_str_PerAnn <> "" And r_str_PerMes <> "" Then
         
            If Trim(Year(Now)) = r_str_PerAnn And (Trim(Month(Now)) = r_str_PerMes Or CInt(Trim(Month(Now))) - 1 = r_str_PerMes) Then
               If (r_str_PerAnn = r_str_AnnAnt And r_str_PerMes = r_str_MesAnt Or (r_str_AnnAnt = "" And r_str_MesAnt = "")) Then
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col2 = r_str_PerAnn
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col3 = r_str_PerMes
                Else
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col2 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr2 = "El año no es válido."
                        
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col3 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr3 = "El mes no es válido."
                End If
            Else
                If CInt(Trim(Month(Now))) = 1 Then
                    If r_str_PerAnn = Trim(Year(Now)) - 1 And (r_str_PerMes = 1 Or r_str_PerMes = 12) Then
                       r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col2 = r_str_PerAnn
                       r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col3 = r_str_PerMes
                    Else
                       r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col2 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value)
                       r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr2 = "El año no es válido."
                        
                       r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col3 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
                       r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr3 = "El mes no es válido."
                    End If
                Else
                    If Trim(Year(Now)) <> r_str_PerAnn Then
                        r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col2 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value)
                        r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr2 = "El año no es válido."
                    Else
                        r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col2 = r_str_PerAnn
                    End If
                    If CInt(Trim(Month(Now))) <> CInt(r_str_PerMes) And CInt(Trim(Month(Now))) - 1 <> CInt(r_str_PerMes) Then
                        r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col3 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
                        r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr3 = "El mes no es válido."
                    Else
                           r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col3 = r_str_PerMes
                       End If
                   End If
               End If
               If r_str_AnnAnt = "" And r_str_MesAnt = "" Then
                   r_str_AnnAnt = r_str_PerAnn
                   r_str_MesAnt = r_str_PerMes
               End If
            Else
               If r_str_PerAnn = "" Then
                   r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col2 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value)
                   r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr2 = "El año no es válido."
               Else
                   r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col2 = r_str_PerAnn
               End If
               If r_str_PerMes = "" Then
                   r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col3 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
                   r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr3 = "El mes no es válido."
               Else
                   r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col3 = r_str_PerMes
               End If
            End If
             
            'Valida Nro.Libro
            If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value) = "" And r_int_RefAnt = r_int_NumRef Then
                r_int_NroLib = r_int_NroLib
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col4 = r_int_NroLib
            Else
                If r_int_RefAnt <> r_int_NumRef And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value) = "" Then
                   r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col4 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)
                   r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr4 = "Número de Libro Inválido."
                Else
                   If IsNumeric(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)) And fs_ValidarNroLibro(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)) = True Then
                       r_int_NroLib = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)
                       r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col4 = r_int_NroLib
                   Else
                       r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col4 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)
                       r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr4 = "Número de Libro Inválido."
                   End If
                End If
            End If
             
             'Valida Fecha de Comprobante
            If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value) = "" And r_int_RefAnt = r_int_NumRef Then
                r_dat_FecCom = r_dat_FecCom
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col5 = r_dat_FecCom
            Else
                If IsDate(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value)) And _
                   Year(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value)) = r_str_PerAnn And _
                   Month(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value)) = r_str_PerMes Then
                   
                    r_dat_FecCom = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col5 = r_dat_FecCom
                Else
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col5 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr5 = "Fecha Inválida."
                End If
            End If
             
             'Valida Tipo Operación
            If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value) = "" And r_int_RefAnt = r_int_NumRef Then
                r_str_TipOpe = r_str_TipOpe
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col6 = r_str_TipOpe
            Else
                If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value)) And fs_ValidarTipOpe(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value)) = True Then
                    r_str_TipOpe = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col6 = r_str_TipOpe
                Else
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col6 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr6 = "Tipo de Operación Inválida."
                End If
            End If
             
            'Valida Moneda
            If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value) = "" And r_int_RefAnt = r_int_NumRef Then
                r_str_TipMon = r_str_TipMon
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col7 = r_str_TipMon
            Else
                If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value)) And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value) <> "" Then
                    r_str_TipMon = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col7 = r_str_TipMon
                Else
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col7 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr7 = "Tipo de Moneda Inválida."
                End If
            End If
             
             'Valida Tipo de Cambio
            If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value) = "" And r_int_RefAnt = r_int_NumRef Then
                r_dbl_TipCam = r_dbl_TipCam
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col8 = r_dbl_TipCam
            Else
                If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value) <> 0 And r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value <> "" Then
                    r_dbl_TipCam = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col8 = r_dbl_TipCam
                Else
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col8 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr8 = "Tipo de Cambio Inválido."
                End If
            End If
            
            'Valida Glosa Cabecera
            If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value) = "" And r_int_RefAnt = r_int_NumRef Then
                r_str_GloCab = r_str_GloCab
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col9 = r_str_GloCab
            Else
                If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value)) And Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value) <> "" Then
                    r_str_GloCab = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col9 = r_str_GloCab
                Else
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col9 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value)
                    r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr9 = "No tiene Glosa la cabecera."
                End If
            End If
            
            'Valida Código Cuenta
            If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value)) And fs_ValidarCtaCtb(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value)) = True Then
                r_str_CtaCtb = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value)
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col10 = r_str_CtaCtb
            Else
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col10 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value)
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr10 = "Cuenta Contable inválida."
            End If
            
            'Valida Glosa Detalle
            If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 11).Value)) Then
                r_str_GloDet = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 11).Value)
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col11 = r_str_GloDet
            Else
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col11 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 11).Value)
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr11 = "No tiene Glosa el Detalle."
            End If
            
            'Valida Fecha Contable
            If IsDate(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 12).Value)) And _
                Year(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 12).Value)) = r_str_PerAnn And _
                Month(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 12).Value)) = r_str_PerMes Then
                
                r_dat_FecCtb = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 12).Value)
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col12 = r_dat_FecCtb
            Else
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col12 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 12).Value)
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr12 = "Fecha Inválida."
            End If
            
            'Valida Flag Debe/Haber
            If Not IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 13).Value)) Then
                r_str_DebHab = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 13).Value)
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col13 = r_str_DebHab
            Else
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col13 = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 13).Value)
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr13 = "No tiene Flag de Debe/Haber."
            End If
            
            'Valida Monto M.N.
            If IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 14).Value)) Or Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 14).Value) = "" Then
                r_dbl_MtoSol = 0
            Else
                r_dbl_MtoSol = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 14).Value)
            End If
            r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col14 = r_dbl_MtoSol
            
            'Valida Monto M.E.
            If IsNull(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 15).Value)) Or Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 15).Value) = "" Then
                r_dbl_MtoDol = 0
            Else
                r_dbl_MtoDol = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 15).Value)
            End If
            r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_Col15 = r_dbl_MtoDol
            
            If r_int_RefAnt = 0 Then
                If r_str_DebHab = "D" Then
                    r_dbl_DebSol = r_dbl_DebSol + r_dbl_MtoSol
                    r_dbl_DebDol = r_dbl_DebDol + r_dbl_MtoDol
                ElseIf r_str_DebHab = "H" Then
                    r_dbl_HabSol = r_dbl_HabSol + r_dbl_MtoSol
                    r_dbl_HabDol = r_dbl_HabDol + r_dbl_MtoDol
                End If
            Else
                If r_int_RefAnt = r_int_NumRef Then
                    If r_str_DebHab = "D" Then
                        r_dbl_DebSol = r_dbl_DebSol + r_dbl_MtoSol
                        r_dbl_DebDol = r_dbl_DebDol + r_dbl_MtoDol
                    ElseIf r_str_DebHab = "H" Then
                        r_dbl_HabSol = r_dbl_HabSol + r_dbl_MtoSol
                        r_dbl_HabDol = r_dbl_HabDol + r_dbl_MtoDol
                    End If
                Else
                    If r_int_RefAnt <> r_int_NumRef Then
                        If r_int_RefAnt <> 0 And r_int_RefAnt <> r_int_NumRef Then
                            If CDbl(Trim(r_dbl_DebSol)) <> CDbl(Trim(r_dbl_HabSol)) Then
                                r_arr_Matriz(UBound(r_arr_Matriz) - 1).CarAsiCtb_ColErr14 = r_arr_Matriz(UBound(r_arr_Matriz) - 1).CarAsiCtb_ColErr14 & vbNewLine & " Monto en Soles del Debe no es igual al Haber."
                            End If
                            If CDbl(Trim(r_dbl_DebDol)) <> CDbl(Trim(r_dbl_HabDol)) Then
                                r_arr_Matriz(UBound(r_arr_Matriz) - 1).CarAsiCtb_ColErr15 = r_arr_Matriz(UBound(r_arr_Matriz) - 1).CarAsiCtb_ColErr15 & vbNewLine & " Monto en Dólares del Debe no es igual al Haber."
                            End If
                        End If
                   
                        r_dbl_DebSol = 0: r_dbl_DebDol = 0
                        r_dbl_HabSol = 0: r_dbl_HabDol = 0
                    End If
                    If r_str_DebHab = "D" Then
                        r_dbl_DebSol = r_dbl_DebSol + r_dbl_MtoSol
                        r_dbl_DebDol = r_dbl_DebDol + r_dbl_MtoDol
                    ElseIf r_str_DebHab = "H" Then
                        r_dbl_HabSol = r_dbl_HabSol + r_dbl_MtoSol
                        r_dbl_HabDol = r_dbl_HabDol + r_dbl_MtoDol
                    End If
                End If
            End If

            If r_int_RefAnt <> r_int_NumRef And r_int_RefAnt <> 0 Then
                If UBound(r_arr_MatVal) = 0 Then
                    ReDim Preserve r_arr_MatVal(UBound(r_arr_MatVal) + 1)
                    r_arr_MatVal(UBound(r_arr_MatVal)).VerAsiCtb_Col1 = r_int_RefAnt
                Else
                    For r_int_ConAux = 1 To UBound(r_arr_MatVal)
                        If r_arr_MatVal(r_int_ConAux).VerAsiCtb_Col1 <> r_int_RefAnt Then
                            r_bol_FlgExi = False
                        Else
                            r_bol_FlgExi = True
                            GoTo Saltar
                        End If
                    Next r_int_ConAux
                End If
Saltar:
                If r_bol_FlgExi = False Then
                    ReDim Preserve r_arr_MatVal(UBound(r_arr_MatVal) + 1)
                    r_arr_MatVal(UBound(r_arr_MatVal)).VerAsiCtb_Col1 = r_int_RefAnt
                Else
                    MsgBox "Los números de referencia deben ser consecutivos.", vbInformation, modgen_g_str_NomPlt
                    GoTo Salir
                End If
            End If
            
            r_int_RefAnt = r_int_NumRef
            r_int_FilExc = r_int_FilExc + 1
        Loop
       
        l_int_NumReg = r_int_FilExc
        If r_int_RefAnt <> 0 Then
            If CDbl(Trim(r_dbl_DebSol)) <> CDbl(Trim(r_dbl_HabSol)) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr14 = r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr14 & vbNewLine & " Monto en Soles del Debe no es igual al Haber."
            End If
            If CDbl(Trim(r_dbl_DebDol)) <> CDbl(Trim(r_dbl_HabDol)) Then
                  r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr15 = r_arr_Matriz(UBound(r_arr_Matriz)).CarAsiCtb_ColErr15 & vbNewLine & " Monto en Dólares del Debe no es igual al Haber."
            End If
        End If
        
        r_dbl_DebSol = 0: r_dbl_DebDol = 0
        r_dbl_HabSol = 0: r_dbl_HabDol = 0
        grd_Listad.Redraw = False
        
        'Cabecera de la Grilla
        grd_Listad.Rows = grd_Listad.Rows + 2
        grd_Listad.FixedRows = 1
        grd_Listad.Rows = grd_Listad.Rows - 1
        grd_Listad.Row = 0
            
        grd_Listad.Col = 0:        grd_Listad.Text = "Nro.Ref.":    grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 1:        grd_Listad.Text = "Año":         grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 2:        grd_Listad.Text = "Mes":         grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 3:        grd_Listad.Text = "Nro.Lib.":    grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 4:        grd_Listad.Text = "Fec.Comp.":   grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 5:        grd_Listad.Text = "Tip.Nota":    grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 6:        grd_Listad.Text = "Cod.Mon.":    grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 7:        grd_Listad.Text = "T.C.":        grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 8:        grd_Listad.Text = "Glosa_Cab":   grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 9:        grd_Listad.Text = "Cta.Ctb.":    grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 10:       grd_Listad.Text = "Glosa_Det":   grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 11:       grd_Listad.Text = "Fec.Ctb.":    grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 12:       grd_Listad.Text = "Deb/Hab":     grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 13:       grd_Listad.Text = "Imp.Sol.":    grd_Listad.CellAlignment = flexAlignCenterCenter
        grd_Listad.Col = 14:       grd_Listad.Text = "Imp.Dol.":    grd_Listad.CellAlignment = flexAlignCenterCenter
        
       For r_int_Contad = 1 To UBound(r_arr_Matriz)
           grd_Listad.Rows = grd_Listad.Rows + 1
           grd_Listad.TextMatrix(r_int_Contad, 0) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr1, r_int_Contad, 0, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 1) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col2
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr2, r_int_Contad, 1, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 2) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col3
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr3, r_int_Contad, 2, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 3) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col4
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr4, r_int_Contad, 3, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 4) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col5
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr5, r_int_Contad, 4, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 5) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col6
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr6, r_int_Contad, 5, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 6) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col7
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr7, r_int_Contad, 6, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 7) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col8
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr8, r_int_Contad, 7, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 8) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col9
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr9, r_int_Contad, 8, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 9) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col10
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr10, r_int_Contad, 9, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 10) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col11
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr11, r_int_Contad, 10, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 11) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col12
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr12, r_int_Contad, 11, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 12) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col13
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr13, r_int_Contad, 12, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 13) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col14
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr14, 0, 13, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
           grd_Listad.TextMatrix(r_int_Contad, 14) = r_arr_Matriz(r_int_Contad).CarAsiCtb_Col15
           Call gs_ColorearCelda(r_arr_Matriz(r_int_Contad).CarAsiCtb_ColErr15, 0, 14, r_arr_Matriz(), r_arr_Matriz(r_int_Contad).CarAsiCtb_Col1)
           
       Next r_int_Contad
       grd_Listad.Redraw = True
       fs_Carga_ArchivoAsiento = True
       
       grd_Listad.Row = 0
       Call gs_RefrescaGrid(grd_Listad)
   End If
   
Salir:
   r_obj_Excel.Quit
   Set r_obj_Excel = Nothing
End Function

Private Sub gs_ColorearCelda(ByVal r_str_DesErr As String, ByVal r_int_NumFil As Integer, ByVal r_int_NumCol As Integer, ByRef r_arr_Mat() As g_tpo_CarAsiCtb, ByVal r_int_NumRefer As Integer)
Dim r_int_ConAux   As Integer

   If Not IsNull(r_str_DesErr) And r_str_DesErr <> "" Then
      If r_int_NumFil = 0 Then
          For r_int_ConAux = 1 To UBound(r_arr_Mat)
              If r_arr_Mat(r_int_ConAux).CarAsiCtb_Col1 = r_int_NumRefer Then
                  grd_Listad.Row = r_int_ConAux
                  grd_Listad.Col = r_int_NumCol
                  grd_Listad.CellBackColor = &H8080FF
                  l_bol_FlgErr = True
              End If
          Next r_int_ConAux
      Else
          If r_arr_Mat(r_int_NumFil).CarAsiCtb_Col1 = r_int_NumRefer Then
              grd_Listad.Row = r_int_NumFil
              grd_Listad.Col = r_int_NumCol
              grd_Listad.CellBackColor = &H8080FF
              l_bol_FlgErr = True
          End If
      End If
    End If
End Sub

Private Function fs_ValidarNroLibro(ByVal r_int_NumLib As Integer) As Boolean
   fs_ValidarNroLibro = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT NRO_LIBRO, DESC_LIBRO FROM CNTBL_LIBRO WHERE NRO_LIBRO = " & r_int_NumLib & ""
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
       g_rst_Genera.Close
       Set g_rst_Genera = Nothing
       Exit Function
   Else
       fs_ValidarNroLibro = True
   End If
End Function

Private Function fs_ValidarTipOpe(ByVal r_str_TipOper As String) As Boolean
   fs_ValidarTipOpe = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT TIPO_NOTA, DESCRIPCION FROM TIPO_NOTA_CNTBL WHERE TIPO_NOTA = '" & r_str_TipOper & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
       g_rst_Genera.Close
       Set g_rst_Genera = Nothing
       Exit Function
   Else
       fs_ValidarTipOpe = True
   End If
End Function

Private Function fs_ValidarCtaCtb(ByVal r_str_CtaCon As String) As Boolean
   fs_ValidarCtaCtb = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT TRIM(CNTA_CTBL) CNTA_CTBL, TRIM(DESC_CNTA) DESC_CNTA "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_CNTA "
   g_str_Parame = g_str_Parame & "  WHERE LENGTH (TRIM(CNTA_CTBL)) = 12 AND CNTA_CTBL = '" & r_str_CtaCon & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
       g_rst_Genera.Close
       Set g_rst_Genera = Nothing
       Exit Function
   Else
       fs_ValidarCtaCtb = True
   End If
End Function

Private Function fs_Actualiza_AsientoContable() As Boolean
Dim r_str_Origen        As String
Dim r_int_Contad        As Integer
Dim r_int_ConAux        As Integer
Dim r_int_NumRef        As Integer
Dim r_int_RefAnt        As Integer

'Cabecera
Dim r_str_PerAnn        As String
Dim r_str_PerMes        As String
Dim r_int_NroLib        As Integer
Dim r_lng_NumAsi        As Long
Dim r_dat_FecCom        As Date
Dim r_str_TipOpe        As String
Dim r_str_TipMon        As String
Dim r_dbl_TipCam        As Double
Dim r_str_GloCab        As String

'Detalle
Dim r_str_CtaCtb        As String
Dim r_str_GloDet        As String
Dim r_dat_FecCtb        As Date
Dim r_str_DebHab        As String
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_int_NumItm        As Integer

   Screen.MousePointer = 11
   r_str_Origen = "LM"
   moddat_g_int_FlgGrb = 1
   l_int_RegTot = 0
   l_int_RegPro = 0
   l_int_RegSPr = 0
   l_int_RegErr = 0
   l_bol_FlgErr = False
    
   fs_Actualiza_AsientoContable = False
   grd_Listad.Redraw = False
   
   ReDim l_arr_LogPro(0)
   ReDim l_arr_LogPro(1)
    
   l_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   l_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   l_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   l_arr_LogPro(1).LogPro_NumErr = 0
        
   'Total de Asientos del archivo
   For r_int_Contad = 1 To grd_Listad.Rows - 1
       r_int_NumRef = Trim(grd_Listad.TextMatrix(r_int_Contad, 0))
       If r_int_RefAnt <> r_int_NumRef Then
           l_int_RegTot = l_int_RegTot + 1
           r_int_RefAnt = r_int_NumRef
       End If
   Next r_int_Contad
   
   r_int_RefAnt = 0
   r_int_NumRef = 0
    
   For r_int_Contad = 1 To grd_Listad.Rows - 1
       r_int_NumRef = Trim(grd_Listad.TextMatrix(r_int_Contad, 0))
       
       If Trim(grd_Listad.TextMatrix(r_int_Contad, 1)) = "" Then
           r_str_PerAnn = r_str_PerAnn
       Else
           r_str_PerAnn = Trim(grd_Listad.TextMatrix(r_int_Contad, 1))
       End If
       If Trim(grd_Listad.TextMatrix(r_int_Contad, 2)) = "" Then
           r_str_PerMes = r_str_PerMes
       Else
           r_str_PerMes = Trim(grd_Listad.TextMatrix(r_int_Contad, 2))
       End If
       If Trim(grd_Listad.TextMatrix(r_int_Contad, 3)) = "" Then
           r_int_NroLib = r_int_NroLib
       Else
           r_int_NroLib = Trim(grd_Listad.TextMatrix(r_int_Contad, 3))
       End If
       If Trim(grd_Listad.TextMatrix(r_int_Contad, 4)) = "" Then
           r_dat_FecCom = r_dat_FecCom
       Else
           r_dat_FecCom = Trim(grd_Listad.TextMatrix(r_int_Contad, 4))
       End If
       If Trim(grd_Listad.TextMatrix(r_int_Contad, 5)) = "" Then
           r_str_TipOpe = r_str_TipOpe
       Else
           r_str_TipOpe = Trim(grd_Listad.TextMatrix(r_int_Contad, 5))
       End If
       If Trim(grd_Listad.TextMatrix(r_int_Contad, 6)) = "" Then
           r_str_TipMon = r_str_TipMon
       Else
           r_str_TipMon = Trim(grd_Listad.TextMatrix(r_int_Contad, 6))
       End If
       If Trim(grd_Listad.TextMatrix(r_int_Contad, 7)) = "" Then
           r_dbl_TipCam = r_dbl_TipCam
       Else
           r_dbl_TipCam = Trim(grd_Listad.TextMatrix(r_int_Contad, 7))
       End If
       If Trim(grd_Listad.TextMatrix(r_int_Contad, 8)) = "" Then
           r_str_GloCab = r_str_GloCab
       Else
           r_str_GloCab = Trim(grd_Listad.TextMatrix(r_int_Contad, 8))
       End If
       r_str_CtaCtb = Trim(grd_Listad.TextMatrix(r_int_Contad, 9))
       r_str_GloDet = Trim(grd_Listad.TextMatrix(r_int_Contad, 10))
       r_dat_FecCtb = Trim(grd_Listad.TextMatrix(r_int_Contad, 11))
       r_str_DebHab = Trim(grd_Listad.TextMatrix(r_int_Contad, 12))
       r_dbl_MtoSol = Trim(grd_Listad.TextMatrix(r_int_Contad, 13))
       r_dbl_MtoDol = Trim(grd_Listad.TextMatrix(r_int_Contad, 14))

      'Obtener Número de Asiento
       If r_int_RefAnt <> r_int_NumRef Then
           r_int_NumItm = 0
           r_lng_NumAsi = modprc_ff_NumAsi(l_arr_LogPro, r_str_PerAnn, r_str_PerMes, r_str_Origen, r_int_NroLib)
           
           'Grabando Cabecera de Asiento
           moddat_g_int_FlgGOK = False
           moddat_g_int_CntErr = 0
       
           'Datos Principales
           Do While moddat_g_int_FlgGOK = False
             g_str_Parame = "USP_INGRESO_CNTBL_ASIENTO_1 ("
             g_str_Parame = g_str_Parame & "'" & r_str_Origen & "', "
             g_str_Parame = g_str_Parame & r_str_PerAnn & ", "
             g_str_Parame = g_str_Parame & r_str_PerMes & ", "
             g_str_Parame = g_str_Parame & CInt(r_int_NroLib) & ", "
             g_str_Parame = g_str_Parame & r_lng_NumAsi & ", "
             g_str_Parame = g_str_Parame & "'" & CStr(Trim(r_str_TipMon)) & "', "
             g_str_Parame = g_str_Parame & CDbl(r_dbl_TipCam) & ", "
             g_str_Parame = g_str_Parame & "'" & CStr(Trim(r_str_TipOpe)) & "', "
             g_str_Parame = g_str_Parame & "'" & r_str_GloCab & "', "
             g_str_Parame = g_str_Parame & "'" & Format(CDate(r_dat_FecCom), "dd/mm/yyyy") & "', "
             g_str_Parame = g_str_Parame & "'" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & "', "
             g_str_Parame = g_str_Parame & "'" & LCase(Mid(modgen_g_str_CodUsu, 1, 5)) & "', "
             g_str_Parame = g_str_Parame & "'1', "
             If moddat_g_int_FlgGrb = 1 Then
                g_str_Parame = g_str_Parame & 0 & ", "
                g_str_Parame = g_str_Parame & 0 & ", "
                g_str_Parame = g_str_Parame & 0 & ", "
                g_str_Parame = g_str_Parame & 0 & ", "
             End If
             g_str_Parame = g_str_Parame & CInt(moddat_g_int_FlgGrb) & ") "
           
             'Datos de Auditoria
             'g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
             'g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
             'g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
             'g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
             'g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ")"
             
             If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                moddat_g_int_CntErr = moddat_g_int_CntErr + 1
                MsgBox g_str_Parame, vbCritical, "Cadena Error"
             Else
                moddat_g_int_FlgGOK = True
             End If
           
             If moddat_g_int_CntErr = 6 Then
                If MsgBox("No se pudo completar el procedimiento USP_INGRESO_CNTBL_ASIENTO. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                   If r_int_RefAnt <> r_int_NumRef Then
                       l_bol_FlgErr = True
                       l_int_RegPro = r_int_ConAux
                       l_int_RegErr = 1
                       l_str_MsjRef = r_int_NumRef
                       l_int_RegSPr = l_int_RegTot - l_int_RegPro - l_int_RegErr
                       l_str_MsjRef = IIf(l_bol_FlgErr = True, "NumRef: " & l_str_MsjRef, "")
                   End If
                   grd_Listad.Redraw = True
                   Screen.MousePointer = 0
                   Exit Function
                Else
                   moddat_g_int_CntErr = 0
                End If
             End If
           Loop
       End If
       
       'Ingresando detalle del asiento
       'Grabando en BD
       moddat_g_int_FlgGOK = False
       moddat_g_int_CntErr = 0
       
       Do While moddat_g_int_FlgGOK = False
       
           g_str_Parame = "USP_INGRESO_CNTBL_ASI_DET_1 ("
           g_str_Parame = g_str_Parame & "'" & r_str_Origen & "', "
           g_str_Parame = g_str_Parame & CStr(r_str_PerAnn) & ", "
           g_str_Parame = g_str_Parame & CStr(r_str_PerMes) & ", "
           g_str_Parame = g_str_Parame & CInt(r_int_NroLib) & ", "
           g_str_Parame = g_str_Parame & CLng(r_lng_NumAsi) & ", "
           g_str_Parame = g_str_Parame & CInt(r_int_NumItm) + 1 & ", "
           
           'Datos de Linea
           g_str_Parame = g_str_Parame & "'" & r_str_CtaCtb & "',"
           g_str_Parame = g_str_Parame & "'" & CDate(r_dat_FecCtb) & "', "
           g_str_Parame = g_str_Parame & "'" & Trim(Mid(Trim(r_str_GloDet), 1, 60)) & "',"
           g_str_Parame = g_str_Parame & "'" & CStr(r_str_DebHab) & "',"
           g_str_Parame = g_str_Parame & CStr(r_dbl_MtoSol) & ", "
           g_str_Parame = g_str_Parame & CStr(r_dbl_MtoDol) & ","
           
           g_str_Parame = g_str_Parame & CInt(moddat_g_int_FlgGrb) & ") "
           
           'Datos de Auditoria
           'g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
           'g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
           'g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
           'g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
           
           If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
              moddat_g_int_CntErr = moddat_g_int_CntErr + 1
           Else
              moddat_g_int_FlgGOK = True
           End If
           
           If moddat_g_int_CntErr = 6 Then
              If MsgBox("No se pudo completar el procedimiento USP_CTB_ASIDET. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
                 If r_int_RefAnt <> r_int_NumRef Then
                    l_bol_FlgErr = True
                    l_int_RegPro = r_int_ConAux
                    l_str_MsjRef = r_int_NumRef
                    l_int_RegErr = 1
                    l_str_MsjRef = r_int_NumRef
                    l_int_RegSPr = l_int_RegTot - l_int_RegPro - l_int_RegErr
                    l_str_MsjRef = IIf(l_bol_FlgErr = True, "NumRef: " & l_str_MsjRef, "")
                 End If
                 grd_Listad.Redraw = True
                 Screen.MousePointer = 0
                 Exit Function
              Else
                 moddat_g_int_CntErr = 0
              End If
           End If
       Loop

       If r_int_RefAnt <> r_int_NumRef Then
           r_int_ConAux = r_int_ConAux + 1
       End If
       r_int_NumItm = r_int_NumItm + 1
       r_int_RefAnt = r_int_NumRef
   Next r_int_Contad
       
   l_int_RegPro = r_int_ConAux
   grd_Listad.Redraw = True
      
   Call gs_RefrescaGrid(grd_Listad)
   Call fs_Limpiar
   Call gs_SetFocus(cmd_BuscaArc)
   
   fs_Actualiza_AsientoContable = True
   l_int_RegErr = IIf(l_bol_FlgErr = True, 1, 0)
   l_int_RegSPr = l_int_RegTot - l_int_RegPro - l_int_RegErr
   l_str_MsjRef = IIf(l_bol_FlgErr = True, "NumRef: " & l_str_MsjRef, "")
End Function

Private Function fs_VerificarCabecera(ByVal r_obj_FicExc As Excel.Application) As Boolean
Dim r_int_NumFil As Integer

   fs_VerificarCabecera = False
   r_int_NumFil = 1
   
   If Trim(r_obj_FicExc.Cells(r_int_NumFil, 1).Value) <> "" Then
       If Trim(r_obj_FicExc.Cells(r_int_NumFil, 1).Value) = "NUMREF" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 2).Value) = "AÑO" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 3).Value) = "MES" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 4).Value) = "NRO_LIBRO" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 5).Value) = "FECHA_COMPR" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 6).Value) = "TIPO_OPER" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 7).Value) = "MONEDA" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 8).Value) = "TIPO_CAMBIO" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 9).Value) = "GLOSA_CABECERA" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 10).Value) = "CNTA_CTBL" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 11).Value) = "GLOSA_DETALLE" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 12).Value) = "FECHA_CNTBL" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 13).Value) = "DEB_HAB" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 14).Value) = "MONTO_MN" Then
           fs_VerificarCabecera = True
       ElseIf Trim(r_obj_FicExc.Cells(r_int_NumFil, 15).Value) = "MONTO_ME" Then
           fs_VerificarCabecera = True
       Else
           fs_VerificarCabecera = False
       End If
   End If
End Function
