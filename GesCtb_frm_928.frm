VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_RptSun_07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   Icon            =   "GesCtb_frm_928.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3405
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   5925
      _Version        =   65536
      _ExtentX        =   10451
      _ExtentY        =   6006
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
         TabIndex        =   9
         Top             =   60
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
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
            Height          =   300
            Left            =   630
            TabIndex        =   10
            Top             =   30
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "SUNAT"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   270
            Left            =   630
            TabIndex        =   11
            Top             =   315
            Width           =   4305
            _Version        =   65536
            _ExtentX        =   7594
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "Facturación Electrónica"
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
            Picture         =   "GesCtb_frm_928.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1275
         Left            =   30
         TabIndex        =   12
         Top             =   2070
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
         _ExtentY        =   2249
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
         Begin VB.CheckBox Chk_FecAct 
            Caption         =   "A la Fecha"
            Height          =   285
            Left            =   1950
            TabIndex        =   3
            Top             =   930
            Width           =   1995
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   150
            Width           =   3465
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1980
            TabIndex        =   2
            Top             =   510
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   -1  'True
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
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Comprobante :"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   210
            Width           =   1395
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha :"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   570
            Width           =   540
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   13
         Top             =   780
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
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
         Begin VB.CommandButton cmd_ConFac 
            Height          =   585
            Left            =   1170
            Picture         =   "GesCtb_frm_928.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Consulta de Documentos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   600
            Picture         =   "GesCtb_frm_928.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Descargar facturas del SFTP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5190
            Picture         =   "GesCtb_frm_928.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Archivo 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_928.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Generar y cargar archivo a SFTP"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   555
         Left            =   30
         TabIndex        =   16
         Top             =   1470
         Width           =   5835
         _Version        =   65536
         _ExtentX        =   10292
         _ExtentY        =   979
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
         Begin VB.ComboBox cmb_TipPro 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   3465
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Proceso:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   180
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frm_RptSun_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents r_chi_sftp  As ChilkatSFtp
Attribute r_chi_sftp.VB_VarHelpID = -1

Dim l_bol_FlgReg           As Boolean
Dim r_int_PerMes           As Integer
Dim r_int_PerAno           As Integer

Dim l_str_NomLOG           As String
Dim l_int_NumLOG           As Integer
Dim l_str_RutaLg           As String
Dim l_str_RutaArc          As String
Dim l_str_FecCar           As String
Dim l_str_RutFacEnt        As String
Dim l_str_RutFacRep        As String
Dim l_str_RutFacAce        As String
Dim l_str_RutFacRec        As String
Dim l_fsobj                As Scripting.FileSystemObject
Dim l_txtStr               As TextStream
Dim l_arr_NumFac()         As moddat_tpo_Genera

' Variables, constantes y declaraciones para el API
Private Type SHFILEOPSTRUCT
    hwnd                   As Long                          ' hWnd del formulario
    wFunc                  As Long                          ' Función a usar: FO_COPY, etc.
    pFrom                  As String                        ' Fichero(s) de origen
    pTo                    As String                        ' Fichero(s) de destino
    ' fFlags    para Windows 2000/XP declararlo como Long
    '           para Windows 9x declararlo como Integer,
    '           aunque también funciona si se declara como Long (al menos en W98)
    'fFlags As Integer                   ' Opciones
    fFlags                 As Long
    fAnyOperationsAborted  As Boolean                       ' Si se ha cancelado
    hNameMappings          As Long
    lpszProgressTitle      As String                        ' Sólo si se usa FOF_SIMPLEPROGRESS
End Type

' Constantes para FileOperation
Private Enum eFO
    FO_COPY = &H2&                                          ' Copiar
    FO_DELETE = &H3&                                        ' Borrar
    FO_MOVE = &H1&                                          ' Mover
    FO_RENAME = &H4&                                        ' Renombrar
    '
    FOF_MULTIDESTFILES = &H1&                               ' Multiples archivos de destino
    FOF_CONFIRMMOUSE = &H2&                                 ' No está implementada
    FOF_SILENT = &H4&                                       ' No mostrar el progreso
    FOF_RENAMEONCOLLISION = &H8&                            ' Cambiar el nombre si el archivo de destino ya existe
    FOF_NOCONFIRMATION = &H10&                              ' No pedir confirmación
    FOF_WANTMAPPINGHANDLE = &H20&
    FOF_ALLOWUNDO = &H40&                                   ' Permitir deshacer
    FOF_FILESONLY = &H80&                                   ' Si se especifica *.*, hacerlo sólo con archivos
    FOF_SIMPLEPROGRESS = &H100&                             ' No mostrar los nombres de los archivos
    FOF_NOCONFIRMMKDIR = &H200&                             ' No confirmar la creación de directorios
    FOF_NOERRORUI = &H400&
    FOF_NOCOPYSECURITYATTRIBS = &H800&
End Enum

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
    (lpFileOp As SHFILEOPSTRUCT) As Long

Private Function fs_Obtener_Codigo(ByVal p_TipDoc As String, ByRef p_CodIte As Long, ByRef p_SerFac As Integer, ByRef p_NumFac As Long)

Dim r_str_Parame           As String
Dim r_rst_Codigo           As ADODB.Recordset
Dim r_int_InsUpd           As Integer
   
   p_CodIte = 0
   p_SerFac = 0
   p_NumFac = 0
   r_int_InsUpd = 0
   
   
   'Código Máximo de CNTBL_DOCELE
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT MAX(DOCELE_CODIGO) AS CODIGO "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_DOCELE "
      
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

   'Número de Serie
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT NVL(DOCELE_NUMSER,0) AS SERIE "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_FOLIOS_DOCELE "
   r_str_Parame = r_str_Parame & "  WHERE DOCELE_TIPDOC = '" & p_TipDoc & "' "
   
   If cmb_TipPro.ItemData(cmb_TipPro.ListIndex) = 9 Then
      r_str_Parame = r_str_Parame & "    AND DOCELE_TIPPRO IS NULL"
   Else
      r_str_Parame = r_str_Parame & "    AND DOCELE_TIPPRO = " & cmb_TipPro.ItemData(cmb_TipPro.ListIndex) & ""
   End If
  
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para obtener codigo en CNTBL_FOLIOS_DOCELE, procedimiento: fs_ObtenerCodigo")
      Exit Function
   End If

   If r_rst_Codigo.BOF And r_rst_Codigo.EOF Then
      p_SerFac = 1
   End If
   
   If Not (r_rst_Codigo.BOF And r_rst_Codigo.EOF) Then
      r_rst_Codigo.MoveFirst
   
      If IsNull(r_rst_Codigo!SERIE) Then
         p_SerFac = 1
      Else
         p_SerFac = r_rst_Codigo!SERIE
      End If
   End If
   
   r_rst_Codigo.Close
   Set r_rst_Codigo = Nothing
   
   'Número de Factura
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT DOCELE_NUMCOR AS CORRELATIVO "
   r_str_Parame = r_str_Parame & "   FROM CNTBL_FOLIOS_DOCELE "
   r_str_Parame = r_str_Parame & "  WHERE DOCELE_TIPDOC = '" & p_TipDoc & "'"
   r_str_Parame = r_str_Parame & "    AND DOCELE_NUMSER = '" & Format(p_SerFac, "000") & "'"
   
   If cmb_TipPro.ItemData(cmb_TipPro.ListIndex) = 9 Then
      r_str_Parame = r_str_Parame & " AND DOCELE_TIPPRO IS NULL"                                               'CUOTAS
   Else
      r_str_Parame = r_str_Parame & " AND DOCELE_TIPPRO = " & cmb_TipPro.ItemData(cmb_TipPro.ListIndex) & ""
   End If
      
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para obtener codigo en CNTBL_FOLIOS_DOCELE, procedimiento: fs_ObtenerCodigo")
      Exit Function
   End If
   
   If r_rst_Codigo.BOF And r_rst_Codigo.EOF Then
      p_NumFac = 0
   End If
   
   If Not (r_rst_Codigo.BOF And r_rst_Codigo.EOF) Then
      r_rst_Codigo.MoveFirst
      If IsNull(r_rst_Codigo!CORRELATIVO) Then
         p_NumFac = 0
      Else
         p_NumFac = r_rst_Codigo!CORRELATIVO
      End If
      r_int_InsUpd = 1
   End If
   
   p_NumFac = p_NumFac + 1
    
   r_rst_Codigo.Close
   Set r_rst_Codigo = Nothing
  
   If p_NumFac = 99999999 Then
      p_NumFac = 0
      p_SerFac = p_SerFac + 1
   End If

   'Actualizando Folio
   If r_int_InsUpd = 0 Then
      'Insert
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " INSERT INTO CNTBL_FOLIOS_DOCELE ("
      r_str_Parame = r_str_Parame & "           DOCELE_TIPDOC, "
      r_str_Parame = r_str_Parame & "           DOCELE_NUMSER, "
      r_str_Parame = r_str_Parame & "           DOCELE_NUMCOR, "
      r_str_Parame = r_str_Parame & "           SEGUSUCRE, "
      r_str_Parame = r_str_Parame & "           SEGFECCRE, "
      r_str_Parame = r_str_Parame & "           SEGHORCRE, "
      r_str_Parame = r_str_Parame & "           SEGPLTCRE, "
      r_str_Parame = r_str_Parame & "           SEGTERCRE, "
      r_str_Parame = r_str_Parame & "           SEGSUCCRE) "
      r_str_Parame = r_str_Parame & " VALUES ("
      r_str_Parame = r_str_Parame & "'" & CStr(p_TipDoc) & "', "
      r_str_Parame = r_str_Parame & "'" & Format(CStr(p_SerFac), "000") & "', "
      r_str_Parame = r_str_Parame & "'" & Format(CStr(p_NumFac), "00000000") & "', "
      r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
      r_str_Parame = r_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
      r_str_Parame = r_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
      r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
      r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
      r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "')"

      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para insertar en CNTBL_FOLIOS_DOCELE, procedimiento: fs_ObtenerCodigo")
         Exit Function
      End If
      
   Else
      'Update
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " UPDATE CNTBL_FOLIOS_DOCELE SET "
      r_str_Parame = r_str_Parame & "        DOCELE_NUMSER = '" & Format(CStr(p_SerFac), "000") & "', "
      r_str_Parame = r_str_Parame & "        DOCELE_NUMCOR = '" & Format(CStr(p_NumFac), "00000000") & "', "
      r_str_Parame = r_str_Parame & "        SEGUSUACT = '" & modgen_g_str_CodUsu & "', "
      r_str_Parame = r_str_Parame & "        SEGFECACT = " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
      r_str_Parame = r_str_Parame & "        SEGHORACT = " & Format(Time, "HHMMSS") & ", "
      r_str_Parame = r_str_Parame & "        SEGPLTACT = '" & UCase(App.EXEName) & "', "
      r_str_Parame = r_str_Parame & "        SEGTERACT = '" & modgen_g_str_NombPC & "', "
      r_str_Parame = r_str_Parame & "        SEGSUCACT = '" & modgen_g_str_CodSuc & "' "
      r_str_Parame = r_str_Parame & "  WHERE "
      r_str_Parame = r_str_Parame & "        DOCELE_TIPDOC = '" & CStr(p_TipDoc) & "' AND "
      
      If cmb_TipPro.ItemData(cmb_TipPro.ListIndex) = 9 Then
         r_str_Parame = r_str_Parame & "     DOCELE_TIPPRO IS NULL"                                               'CUOTAS
      Else
         r_str_Parame = r_str_Parame & "     DOCELE_TIPPRO = " & cmb_TipPro.ItemData(cmb_TipPro.ListIndex) & ""
      End If
          
      If Not gf_EjecutaSQL(r_str_Parame, r_rst_Codigo, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta para actualizar en CNTBL_FOLIOS_DOCELE, procedimiento: fs_ObtenerCodigo")
         Exit Function
      End If
   End If
   
   Set r_rst_Codigo = Nothing
   
End Function
Private Sub fs_Inicia()
   cmb_TipDoc.Clear
   
   cmb_TipDoc.AddItem "- TODOS -"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = 0
   cmb_TipDoc.AddItem "FACTURAS"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = "01"
   cmb_TipDoc.AddItem "BOLETAS"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = "03"
   cmb_TipDoc.AddItem "NOTAS DE CREDITO"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = "07"
   cmb_TipDoc.AddItem "PREPAGOS"
   cmb_TipDoc.ItemData(cmb_TipDoc.NewIndex) = "05"
   
   cmb_TipDoc.ListIndex = 0
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPro, 1, "539")
End Sub
Private Function fs_Cargar_Archivo(p_sRuta, p_sFile) As Boolean

Dim r_str_success    As Long
Dim r_dbl_port       As Long
Dim r_str_HostName   As String
Dim r_str_RutServ    As String
Dim r_key            As New ChilkatSshKey
Dim r_privKey        As String
   
   On Error GoTo MyError
   
   fs_Cargar_Archivo = False
   
   Set r_chi_sftp = New ChilkatSFtp
   
   r_str_success = r_chi_sftp.UnlockComponent("30")
   If (r_str_success <> 1) Then
       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
       Exit Function
   End If
   
   'Set some timeouts, in milliseconds:
   r_chi_sftp.ConnectTimeoutMs = 5000
   r_chi_sftp.IdleTimeoutMs = 10000
   
   '  Connect to the SSH server.
   '  The standard SSH port = 22
   '  The hostname may be a hostname or IP address.
   
   '  Producción:
   '  Sftp.escondatagate.net (puerto 6022)
   '  Calidad:
   '  Sftpqa.escondatagate.net (puerto 3022)
   
   r_str_HostName = " Sftp.escondatagate.net"
   r_dbl_port = 6022
   
'   r_str_HostName = "Sftpqa.escondatagate.net"
'   r_dbl_port = 3022
   
   r_str_success = r_chi_sftp.Connect(r_str_HostName, r_dbl_port)
   If (r_str_success <> 1) Then
       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
       Exit Function
   End If
    
   'clave pública
   r_privKey = r_key.LoadText(moddat_g_str_RutFac & "\" & "id_rsa.ppk")
   
   If (r_key.LastMethodSuccess <> 1) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_key.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
      Exit Function
   End If
   
   r_str_success = r_key.FromOpenSshPrivateKey(r_privKey)
   If (r_str_success <> 1) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_key.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
      Exit Function
   End If
   
   '  r_str_success = r_chi_sftp.AuthenticatePk("micasi02", key)
   
   '  Authenticate with the SSH server.  Chilkat SFTP supports
   '  both password-based authenication as well as public-key
   '  authentication.  This example uses password authenication.
   '  r_str_success = r_chi_sftp.AuthenticatePw("micasi01", "Micasi2018*")
   
   r_str_success = r_chi_sftp.AuthenticatePw("micasi02", "Micasi2018*")
   
   If (r_str_success <> 1) Then
       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
       Exit Function
   End If
   
   '  After authenticating, the SFTP subsystem must be initialized:
   r_str_success = r_chi_sftp.InitializeSftp()
   If (r_str_success <> 1) Then
       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
       Exit Function
   End If
   
   r_str_RutServ = "/WWW/entrada/" & p_sFile

   r_str_success = r_chi_sftp.UploadFileByName(r_str_RutServ, p_sRuta)
   If (r_str_success <> 1) Then
       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
       Exit Function
   End If
   
   fs_Cargar_Archivo = True
   Exit Function
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Cargar_Archivo")
End Function
Private Function fs_NroEnLetras(ByVal curNumero As Double, Optional blnO_Final As Boolean = True) As String

    Dim dblCentavos As Double
    Dim lngContDec As Long
    Dim lngContCent As Long
    Dim lngContMil As Long
    Dim lngContMillon As Long
    Dim strNumLetras As String
    Dim strNumero As Variant
    Dim strDecenas As Variant
    Dim strCentenas As Variant
    Dim blnNegativo As Boolean
    Dim blnPlural As Boolean
                
    If Int(curNumero) = 0# Then
        strNumLetras = "CERO"
    End If
    
    strNumero = Array(vbNullString, "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", _
                   "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", _
                   "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", _
                   "VEINTE")

    strDecenas = Array(vbNullString, vbNullString, "VEINTI", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", _
                    "SETENTA", "OCHENTA", "NOVENTA", "CIEN")

    strCentenas = Array(vbNullString, "CIENTO", "DOSCIENTOS", "TRESCIENTOS", _
                     "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", _
                     "OCHOCIENTOS", "NOVECIENTOS")

    If curNumero < 0# Then
        blnNegativo = True
        curNumero = Abs(curNumero)
    End If

    If Int(curNumero) <> curNumero Then
        dblCentavos = Abs(curNumero - Int(curNumero))
        curNumero = Int(curNumero)
    End If

    Do While curNumero >= 1000000#
        lngContMillon = lngContMillon + 1
        curNumero = curNumero - 1000000#
    Loop

    Do While curNumero >= 1000#
        lngContMil = lngContMil + 1
        curNumero = curNumero - 1000#
    Loop
    
    Do While curNumero >= 100#
        lngContCent = lngContCent + 1
        curNumero = curNumero - 100#
    Loop
    
    If Not (curNumero > 10# And curNumero <= 20#) Then
        Do While curNumero >= 10#
            lngContDec = lngContDec + 1
            curNumero = curNumero - 10#
        Loop
    End If
    
    If lngContMillon > 0 Then
        If lngContMillon >= 1 Then   'si el número es >1000000 usa recursividad
            strNumLetras = fs_NroEnLetras(lngContMillon, False)
            If Not blnPlural Then blnPlural = (lngContMillon > 1)
            lngContMillon = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMillon) & " MILLON" & _
                                                                    IIf(blnPlural, "ES ", " ")
    End If
    
    If lngContMil > 0 Then
        
        If lngContMil = 1 Then   'si el número es >100000 usa recursividad
            strNumLetras = strNumLetras & fs_NroEnLetras(lngContMil, False)
            lngContMil = 0
            
        End If
        If lngContMil > 1 Then   'si el número es >100000 usa recursividad
            strNumLetras = strNumLetras & fs_NroEnLetras(lngContMil, False)
            lngContMil = 0
        End If
        'MsgBox strNumLetras
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMil) & " MIL "
        'MsgBox strNumLetras
    End If
    
    If lngContCent > 0 Then
        If lngContCent = 1 And lngContDec = 0 And curNumero = 0# Then
            strNumLetras = strNumLetras & "CIEN"
        Else
            strNumLetras = strNumLetras & strCentenas(lngContCent) & " "
        End If
    End If
    If lngContDec >= 1 Then
        If lngContDec = 1 Then
            strNumLetras = strNumLetras & strNumero(10)
        Else
            strNumLetras = strNumLetras & strDecenas(lngContDec)
        End If
        
        If lngContDec >= 3 And curNumero > 0# Then
            strNumLetras = strNumLetras & " Y "
        End If
    Else
    'MsgBox "Por Aqui"
        If curNumero >= 0# And curNumero <= 20# Then
            strNumLetras = strNumLetras & strNumero(curNumero)
            If curNumero = 1# And blnO_Final Then
                strNumLetras = strNumLetras & "O"
            End If
            If dblCentavos > 0# Then
            
                strNumLetras = Trim(strNumLetras) & " CON " & Format$(CInt(dblCentavos * 100#), "00") & "/100"
            Else

                'strNumLetras = Trim(strNumLetras) & " CON " & Format$(CInt(dblCentavos * 100#), "00") & "/100"
            End If
            fs_NroEnLetras = strNumLetras
            Exit Function
        End If
    End If
    
    If curNumero > 0# Then
    
        strNumLetras = strNumLetras & strNumero(curNumero)
        If curNumero = 1# And blnO_Final Then
            strNumLetras = strNumLetras & "O"
        End If
    End If
    
    If dblCentavos > 0# Then
        strNumLetras = strNumLetras & " CON " + Format$(CInt(dblCentavos * 100#), "00") & "/100"
    'Else
    End If
    'If dblCentavos = 0# Then
        'MsgBox strNumLetras
        'strNumLetras = strNumLetras & " CON " + Format$(CInt(dblCentavos * 100#), "00") & "/100"
    'End If
    
    fs_NroEnLetras = IIf(blnNegativo, "(" & strNumLetras & ")", strNumLetras)
    
End Function

Private Sub Chk_FecAct_Click()
   If Chk_FecAct.Value = 1 Then
      ipp_FecIni.Text = ""
      ipp_FecIni.Enabled = False
   ElseIf Chk_FecAct.Value = 0 Then
      ipp_FecIni.Enabled = True
      Call fs_Limpia
   End If
   Call gs_SetFocus(Chk_FecAct)
End Sub
Private Sub Chk_FecAct_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Archivo)
  End If
End Sub


Private Sub cmd_Archivo_Click()
Dim r_str_RutLog     As String
   
'   If cmb_TipPro.ListIndex = -1 Then
'      MsgBox "Debe seleccionar Tipo de Proceso.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(cmb_TipPro)
'      Exit Sub
'   End If
   
   If cmb_TipDoc.ListIndex = 0 Then
      MsgBox "Debe seleccionar Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
      
   If MsgBox("¿Está seguro de enviar archivo?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   If Chk_FecAct.Value = 0 Then
      l_str_FecCar = Format(ipp_FecIni.Text, "yyyymmdd")
   Else
      l_str_FecCar = Format(Now, "yyyymmdd")
   End If
     
   'Crear Archivo LOG del Proceso
   l_str_NomLOG = UCase(App.EXEName) & "_C_" & Format(date, "yyyymmdd") & ".LOG"
   l_int_NumLOG = FreeFile
   
   r_str_RutLog = Replace(moddat_g_str_RutFac, "\Fact", "\Logs")
   
   If gf_Existe_Archivo(r_str_RutLog & "\", l_str_NomLOG) Then
      Kill r_str_RutLog & "\" & l_str_NomLOG
      DoEvents
   End If
   
   l_str_RutaLg = r_str_RutLog & "\" & l_str_NomLOG
   Open l_str_RutaLg For Output As l_int_NumLOG
   Close #l_int_NumLOG
   
   Call fs_Escribir_Linea(l_str_RutaLg, "")
   Call fs_Escribir_Linea(l_str_RutaLg, "Proceso           : " & modgen_g_str_NomPlt)
   Call fs_Escribir_Linea(l_str_RutaLg, "Proceso           : " & modgen_g_str_NomPlt)
   Call fs_Escribir_Linea(l_str_RutaLg, "Nombre Ejecutable : " & UCase(App.EXEName))
   Call fs_Escribir_Linea(l_str_RutaLg, "Número Revisión   : " & modgen_g_str_NumRev)
   Call fs_Escribir_Linea(l_str_RutaLg, "Nombre PC         : " & modgen_g_str_NombPC)
   Call fs_Escribir_Linea(l_str_RutaLg, "Origen Datos      : " & moddat_g_str_NomEsq & " - " & moddat_g_str_EntDat)
   Call fs_Escribir_Linea(l_str_RutaLg, "")
   Call fs_Escribir_Linea(l_str_RutaLg, "Inicio Proceso    : " & Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss"))
   Call fs_Escribir_Linea(l_str_RutaLg, "")
       
   'Ruta para almacenar los archivos generados de facturas
   l_str_RutFacEnt = moddat_g_str_RutFac & "\entrada\" 'moddat_g_str_RutLoc
   
   'Crear la Carpeta entrada
   Set l_fsobj = New FileSystemObject
   If l_fsobj.FolderExists(l_str_RutFacEnt) = False Then
      l_fsobj.CreateFolder (l_str_RutFacEnt)
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = "01" Then
      Call fs_Generar_Facturas                                                                     'Ingresando Facturas de Fechas Anteriores a Hoy
      Call fs_Generar_Archivo_01_03(l_str_FecCar, "01")                                            'Generando el archivo UTF8 - No Bom - Facturas (01)
      'Call fs_Generar_Facturas_NUEVO_FORMATO
      
   ElseIf cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = "03" Or cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = "05" Then
      Call fs_Generar_Boletas
      Call fs_Generar_Archivo_01_03(l_str_FecCar, "03")                                           'Generando el archivo UTF8 - No Bom - Boletas (03)
'      Call fs_Generar_Boletas_NUEVO_FORMATO
'      Call fs_Generar_Archivo_01_03_NUEVO_FORMATO(l_str_FecCar, "03")                            'Generando el archivo UTF8 - No Bom - Boletas (03)
'
   ElseIf cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = "07" Then
      Call fs_Generar_Archivo_07(l_str_FecCar, "07")                                               'Generando el archivo UTF8 - No Bom - Notas de Crédito (07)
   End If
   
   'Cerrando Archivo LOG del Proceso
   Call fs_Escribir_Linea(l_str_RutaLg, "")
   Call fs_Escribir_Linea(l_str_RutaLg, "Fecha Proceso     : " & Format(date, "dd/mm/yyyy"))
   Call fs_Escribir_Linea(l_str_RutaLg, "Fin Proceso       : " & Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss"))
   Call fs_Escribir_Linea(l_str_RutaLg, "")
   
   'Para enviar Correo Electrónico
   'Call fs_Envia_CorEle_User
   'Call fs_Envia_CorEle_LOG
   
   MsgBox "Proceso Finalizado.", vbInformation, modgen_g_str_NomPlt
   Screen.MousePointer = 0
    
End Sub
Private Sub fs_Generar_Facturas()
Dim r_lng_Contad     As Long
Dim r_int_SerFac     As Integer
Dim r_lng_NumFac     As Long

   On Error GoTo MyError
   
   Screen.MousePointer = 11
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT  CAMPO_IDE_01, CAMPO_IDE_02  , CAMPO_IDE_03 , CAMPO_IDE_04 , CAMPO_IDE_05   , CAMPO_IDE_06   , CAMPO_IDE_07   , CAMPO_IDE_08 , CAMPO_EMI_01  , CAMPO_EMI_02,"
   g_str_Parame = g_str_Parame & "           CAMPO_EMI_03, CAMPO_EMI_04  , CAMPO_EMI_05 , CAMPO_EMI_06 , CAMPO_EMI_07   , CAMPO_EMI_08   , CAMPO_EMI_09   , CAMPO_EMI_10 , CAMPO_EMI_11  , CAMPO_EMI_12,"
   g_str_Parame = g_str_Parame & "           CAMPO_EMI_13, CAMPO_EMI_14  , CAMPO_EMI_15 , CAMPO_REC_01 , CAMPO_REC_02   , CAMPO_REC_03   , CAMPO_REC_04   , CAMPO_REC_05 , CAMPO_REC_06  , CAMPO_REC_07,"
   g_str_Parame = g_str_Parame & "           CAMPO_REC_08, CAMPO_REC_09  , CAMPO_REC_10 , CAMPO_REC_11 , CAMPO_REC_12   , CAMPO_DRF_01   , CAMPO_DRF_02   , CAMPO_DRF_03 , CAMPO_DRF_04  , CAMPO_DRF_05,"
   g_str_Parame = g_str_Parame & "           CAMPO_DRF_06, CAMPO_CAB_01  , CAMPO_CAB_02 , CAMPO_CAB_03 , CAMPO_CAB_04   , CAMPO_CAB_05   , CAMPO_CAB_06   , CAMPO_CAB_07 , CAMPO_CAB_08  , CAMPO_CAB_09,"
   g_str_Parame = g_str_Parame & "           CAMPO_CAB_10, CAMPO_CAB_11  , CAMPO_CAB_12 , CAMPO_CAB_13 , CAMPO_CAB_14   , CAMPO_CAB_15   , CAMPO_CAB_16   , CAMPO_CAB_17 , CAMPO_CAB_18_1, CAMPO_CAB_18_2,"
   g_str_Parame = g_str_Parame & "           CAMPO_CAB_19, CAMPO_CAB_20  , CAMPO_CAB_21 , CAMPO_CAB_22 , CAMPO_CAB_23   , CAMPO_CAB_24   , CAMPO_CAB_25   , CAMPO_CAB_26 , CAMPO_CAB_27  , CAMPO_DET_01,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_02, CAMPO_DET_03  , CAMPO_DET_04 , "
   g_str_Parame = g_str_Parame & "           MIN(CAMPO_DET_05) AS CAMPO_DET_05, "
   g_str_Parame = g_str_Parame & "           CAMPO_DET_06, CAMPO_DET_07  , CAMPO_DET_08 , CAMPO_DET_09 , CAMPO_DET_10_1 , CAMPO_DET_10_2 , CAMPO_DET_10_3 , CAMPO_DET_11 , CAMPO_DET_12  , CAMPO_DET_13,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_14, CAMPO_DET_15  , CAMPO_DET_16 , CAMPO_DET_17 , CAMPO_DET_18   , CAMPO_DET_19   , CAMPO_DET_20   , CAMPO_DET_21 , CAMPO_DET_22  , CAMPO_DET_23,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_24, CAMPO_DET_25  , CAMPO_DET_26 , CAMPO_DET_27 , CAMPO_DET_28   , CAMPO_DET_29   , CAMPO_DET_30   , CAMPO_DET_31 , CAMPO_DET_32  , CAMPO_DET_33,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_34, CAMPO_DET_35  , CAMPO_DET_36 , CAMPO_DET_37 , CAMPO_DET2_01  , CAMPO_DET2_02  , CAMPO_DET2_03  , CAMPO_DET2_04,"
   g_str_Parame = g_str_Parame & "           MIN(CAMPO_DET2_05) AS CAMPO_DET2_05, "
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_06, CAMPO_DET2_07, CAMPO_DET2_08, CAMPO_DET2_09, CAMPO_DET2_10_1, CAMPO_DET2_10_2, CAMPO_DET2_10_3, CAMPO_DET2_11, CAMPO_DET2_12 , CAMPO_DET2_13,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_14, CAMPO_DET2_15, CAMPO_DET2_16, CAMPO_DET2_17, CAMPO_DET2_18  , CAMPO_DET2_19  , CAMPO_DET2_20  , CAMPO_DET2_21, CAMPO_DET2_22 , CAMPO_DET2_23,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_24, CAMPO_DET2_25, CAMPO_DET2_26, CAMPO_DET2_27, CAMPO_DET2_28  , CAMPO_DET2_29  , CAMPO_DET2_30  , CAMPO_DET2_31, CAMPO_DET2_32 , CAMPO_DET2_33,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_34, CAMPO_DET2_35, CAMPO_DET2_36, CAMPO_DET2_37, CAMPO_ADI_01   , CAMPO_ADI_02   , CAMPO_ADI_03   , CAMPO_ADI_04 , OPERACION     , NUMERO_MOVIMIENTO,"
   g_str_Parame = g_str_Parame & "           FECHA_CANCELACION , SITUACION, FECHA_DEPOSITO "
   g_str_Parame = g_str_Parame & "     FROM ( "
'''
   g_str_Parame = g_str_Parame & "     SELECT 'IDE'                                                                                               AS CAMPO_IDE_01, "
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_IDE_02, "  '001-
   g_str_Parame = g_str_Parame & "            SUBSTR(CAJMOV_FECDEP,1,4) || '-' || SUBSTR(CAJMOV_FECDEP,5,2) || '-' || SUBSTR(CAJMOV_FECDEP,7,2)   AS CAMPO_IDE_03, "
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_04,"
   g_str_Parame = g_str_Parame & "            '01'                                                                                                AS CAMPO_IDE_05,              " '--CATALOGO N°1.- FACTURA
   g_str_Parame = g_str_Parame & "            C.CATSUN_CODIGO                                                                                     AS CAMPO_IDE_06,              " '--CATALOGO N°2.-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_07,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_08,"
   g_str_Parame = g_str_Parame & "            'EMI'                                                                                               AS CAMPO_EMI_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_EMI_02,"  '001-
   g_str_Parame = g_str_Parame & "            '6'                                                                                                 AS CAMPO_EMI_03,              " '--CATALOGO N°6.-
   g_str_Parame = g_str_Parame & "            '20511904162'                                                                                       AS CAMPO_EMI_04,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_05,"
   g_str_Parame = g_str_Parame & "            'EDPYME MICASITA SA'                                                                                AS CAMPO_EMI_06,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_07,              " '--ATALOGO N°13.-UBIGEO
   g_str_Parame = g_str_Parame & "            'AV RIVERA NAVARRETE 645'                                                                           AS CAMPO_EMI_08,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_09,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_10,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_11,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_12,"
   g_str_Parame = g_str_Parame & "            'PE'                                                                                                AS CAMPO_EMI_13,              " '--CATALOGO N°4.-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_14,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_15,"
   g_str_Parame = g_str_Parame & "            'REC'                                                                                               AS CAMPO_REC_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_REC_02,"  '001-
   g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_TIPDOC)                                                                               AS CAMPO_REC_03,              " '--CATALOGO N°6.-
   g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_NUMDOC)                                                                               AS CAMPO_REC_04,"
   g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_APEPAT) || ' ' || TRIM(D.DATGEN_APEMAT) || ' ' || TRIM(D.DATGEN_NOMBRE)               AS CAMPO_REC_05,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_REC_06,"
   g_str_Parame = g_str_Parame & "            TRIM(H.PARDES_DESCRI)                                                                               AS CAMPO_REC_07,"
   g_str_Parame = g_str_Parame & "            TRIM(I.PARDES_DESCRI)                                                                               AS CAMPO_REC_08,"
   g_str_Parame = g_str_Parame & "            TRIM(G.PARDES_DESCRI)                                                                               AS CAMPO_REC_09,"
   g_str_Parame = g_str_Parame & "            'PE'                                                                                                AS CAMPO_REC_10,"
   g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_TELEFO)                                                                               AS CAMPO_REC_11,"
   g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_DIRELE)                                                                               AS CAMPO_REC_12,"
   g_str_Parame = g_str_Parame & "            'DRF'                                                                                               AS CAMPO_DRF_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_DRF_02," '001-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_03,               "
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_04,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_05,               " '--PARA NOTA CREDITO/DEBITO
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_06,"
   g_str_Parame = g_str_Parame & "            'CAB'                                                                                               AS CAMPO_CAB_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_CAB_02," '001-
   g_str_Parame = g_str_Parame & "            '1001'                                                                                              AS CAMPO_CAB_03,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_04,"
   g_str_Parame = g_str_Parame & "            '1002'                                                                                              AS CAMPO_CAB_05,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_06,"
   g_str_Parame = g_str_Parame & "            '1003'                                                                                              AS CAMPO_CAB_07,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            '0.00'                                                                                              AS CAMPO_CAB_08,"
   g_str_Parame = g_str_Parame & "            '1004'                                                                                              AS CAMPO_CAB_09,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            '0.00'                                                                                              AS CAMPO_CAB_10,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_CAB_11,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_12,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_13,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_14,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_15,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_16,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_17,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_CAB_18_1,"
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_18_2,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_19,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_20,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_21,"
   g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_22,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_23,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_24,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_25,"
   g_str_Parame = g_str_Parame & "            '13'                                                                                                AS CAMPO_CAB_26,                " '--CATALOGO N°17.-
   g_str_Parame = g_str_Parame & "            '[1000'                                                                                             AS CAMPO_CAB_27,                " '--CATALOGO N°15.- DETALLE EN LETRAS DEL IMPORTE
   g_str_Parame = g_str_Parame & "            'DET1'                                                                                              AS CAMPO_DET_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_DET_02," '001-
   g_str_Parame = g_str_Parame & "            '001'                                                                                               AS CAMPO_DET_03,                " '-- Número de orden de ítem
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_04,                "
   
   g_str_Parame = g_str_Parame & "            'INTERES ' ||"
   g_str_Parame = g_str_Parame & "            SUBSTR(TO_CHAR(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(HIPCUO_FECVCT,1,4)  AS CAMPO_DET_05," 'CAJMOV_FECDEP
   g_str_Parame = g_str_Parame & "            1.000                                                                                               AS CAMPO_DET_06,"
   g_str_Parame = g_str_Parame & "            'NIU'                                                                                               AS CAMPO_DET_07,                " '--CATALOGO N°3.-
   g_str_Parame = g_str_Parame & "            A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_08,"
   g_str_Parame = g_str_Parame & "            1.000 * A.CAJMOV_INTERE                                                                             AS CAMPO_DET_09,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_DET_10_1,              " '--CATALOGO N°5,7 u 8.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_DET_10_2,"
   g_str_Parame = g_str_Parame & "            '30'                                                                                                AS CAMPO_DET_10_3,"
   g_str_Parame = g_str_Parame & "            A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_11,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_12,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_13,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_14,"
   g_str_Parame = g_str_Parame & "            '84121901'                                                                                          AS CAMPO_DET_15,                " '--CATALOGO N°15.-
   g_str_Parame = g_str_Parame & "            '7004'                                                                                              AS CAMPO_DET_16,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_17,                " '--NRO CONTRATO O PRESTAMO
   g_str_Parame = g_str_Parame & "            '7005'                                                                                              AS CAMPO_DET_18,"
   g_str_Parame = g_str_Parame & "            SUBSTR(J.HIPMAE_FECACT,1,4) || '-' || SUBSTR(J.HIPMAE_FECACT,5,2) || '-' || SUBSTR(J.HIPMAE_FECACT,7,2)"
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET_19,"
   g_str_Parame = g_str_Parame & "            '7001'                                                                                              AS CAMPO_DET_20,"
   g_str_Parame = g_str_Parame & "            '1'                                                                                                 AS CAMPO_DET_21,                " '--CATALOGO N°26.- Si es construcción /adquisición
   g_str_Parame = g_str_Parame & "            '7003'                                                                                              AS CAMPO_DET_22,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_23,"
   g_str_Parame = g_str_Parame & "            '7002'                                                                                              AS CAMPO_DET_24,"
   g_str_Parame = g_str_Parame & "            CASE WHEN L.SOLMAE_PRIVIV = 1 THEN 3  "
   g_str_Parame = g_str_Parame & "                 WHEN L.SOLMAE_PRIVIV = 2 THEN 0 END                                                            AS CAMPO_DET_25,                " '--CATALOGO N°27.- VERIFICAR
   g_str_Parame = g_str_Parame & "            '7007'                                                                                              AS CAMPO_DET_26,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_27,"
   g_str_Parame = g_str_Parame & "            '7006'                                                                                              AS CAMPO_DET_28,"
   g_str_Parame = g_str_Parame & "            '-'/*P.SOLINM_UBIGEO*/                                                                              AS CAMPO_DET_29,"
   g_str_Parame = g_str_Parame & "            '7008'                                                                                              AS CAMPO_DET_30,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_31,"
   g_str_Parame = g_str_Parame & "            '7011'                                                                                              AS CAMPO_DET_32,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_33,"
   g_str_Parame = g_str_Parame & "            '7009'                                                                                              AS CAMPO_DET_34,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_35,"
   g_str_Parame = g_str_Parame & "            '7010'                                                                                              AS CAMPO_DET_36,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_37,"
   g_str_Parame = g_str_Parame & "            'DET2'                                                                                              AS CAMPO_DET2_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_DET2_02," '001-
   g_str_Parame = g_str_Parame & "            '002'                                                                                               AS CAMPO_DET2_03,                " '-- Número de orden de ítem
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_04,                "
   
   g_str_Parame = g_str_Parame & "            'OTROS IMPORTES ' ||"
   g_str_Parame = g_str_Parame & "            SUBSTR(TO_CHAR(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(HIPCUO_FECVCT,1,4)  AS CAMPO_DET2_05," 'CAJMOV_FECDEP
   g_str_Parame = g_str_Parame & "            1.000                                                                                               AS CAMPO_DET2_06,"
   g_str_Parame = g_str_Parame & "            'NIU'                                                                                               AS CAMPO_DET2_07,                " '--CATALOGO N°3.-
   g_str_Parame = g_str_Parame & "            (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)"
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_08,"
   g_str_Parame = g_str_Parame & "            1.000 * (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)"
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_09,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_DET2_10_1,              " '--CATALOGO N°5,7 u 8.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_DET2_10_2,"
   g_str_Parame = g_str_Parame & "            '30'                                                                                                AS CAMPO_DET2_10_3,"
   g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG - A.CAJMOV_INTERE "
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_11,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_12,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_13,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_14,"
   g_str_Parame = g_str_Parame & "            '84121501'                                                                                          AS CAMPO_DET2_15,                " '--CATALOGO N°15.-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_16,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_17,                " '--NRO CONTRATO O PRESTAMO
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_18,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_19,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_20,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_21,                " '--CATALOGO N°26.- Si es construcción /adquisición
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_22,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_23,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_24,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_25,                " '--CATALOGO N°27.- VERIFICAR
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_26,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_27,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_28,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_29,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_30,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_31,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_32,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_33,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_34,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_35,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_36,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_37,"

   g_str_Parame = g_str_Parame & "            'ADI1'                                                                                              AS CAMPO_ADI_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_ADI_02," '001-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_ADI_03,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_ADI_04,"
   g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_NUMOPE)                                                                               AS OPERACION, "
   g_str_Parame = g_str_Parame & "            CAJMOV_NUMMOV                                                                                       AS NUMERO_MOVIMIENTO , "
   g_str_Parame = g_str_Parame & "            HIPMAE_FECCAN                                                                                       AS FECHA_CANCELACION ,"
   g_str_Parame = g_str_Parame & "            HIPMAE_SITUAC                                                                                       AS SITUACION, "
   g_str_Parame = g_str_Parame & "            CAJMOV_FECDEP                                                                                       AS FECHA_DEPOSITO "
   g_str_Parame = g_str_Parame & "       FROM OPE_CAJMOV A "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '204' AND B.PARDES_CODITE = A.CAJMOV_MONPAG "
   g_str_Parame = g_str_Parame & "            INNER JOIN CNTBL_CATSUN C ON C.CATSUN_NROCAT = 2 AND C.CATSUN_DESCRI = TRIM(B.PARDES_DESCRI) "
   g_str_Parame = g_str_Parame & "            INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = A.CAJMOV_TIPDOC AND D.DATGEN_NUMDOC = A.CAJMOV_NUMDOC "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 201 AND E.PARDES_CODITE = D.DATGEN_TIPVIA "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 202 AND F.PARDES_CODITE = D.DATGEN_TIPZON "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES G ON G.PARDES_CODGRP = 101 AND G.PARDES_CODITE = D.DATGEN_UBIGEO "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 101 AND H.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,2)||'0000' "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES I ON I.PARDES_CODGRP = 101 AND I.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,4)||'00' "
   g_str_Parame = g_str_Parame & "            INNER JOIN CRE_HIPMAE J ON J.HIPMAE_NUMOPE = A.CAJMOV_NUMOPE "
   g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_HIPGAR K ON K.HIPGAR_NUMOPE = J.HIPMAE_NUMOPE AND K.HIPGAR_BIEGAR = 1 "
   g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_SOLMAE L ON L.SOLMAE_NUMERO = J.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_SOLINM P ON P.SOLINM_NUMSOL = J.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES Q ON Q.PARDES_CODGRP = 201 AND Q.PARDES_CODITE = P.SOLINM_TIPVIA "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES R ON R.PARDES_CODGRP = 202 AND R.PARDES_CODITE = P.SOLINM_TIPZON "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES S ON S.PARDES_CODGRP = 101 AND S.PARDES_CODITE = P.SOLINM_UBIGEO "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES T ON T.PARDES_CODGRP = 101 AND T.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,4)||'00' "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES U ON U.PARDES_CODGRP = 101 AND U.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,2)||'0000' "
   
   g_str_Parame = g_str_Parame & "            INNER JOIN CRE_HIPPAG V ON V.HIPPAG_NUMOPE = A.CAJMOV_NUMOPE AND V.HIPPAG_FECPAG = A.CAJMOV_FECDEP AND V.HIPPAG_NUMMOV = A.CAJMOV_NUMMOV"
   g_str_Parame = g_str_Parame & "            INNER JOIN CRE_HIPCUO W ON W.HIPCUO_NUMOPE = V.HIPPAG_NUMOPE AND W.HIPCUO_TIPCRO = 1 AND W.HIPCUO_NUMCUO = V.HIPPAG_NUMCUO"
               
   g_str_Parame = g_str_Parame & "      WHERE CAJMOV_SUCMOV IS NOT NULL "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_USUMOV IS NOT NULL "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_FECMOV > 0"
   g_str_Parame = g_str_Parame & "        AND CAJMOV_NUMMOV > 0 "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_CODBAN IS NOT NULL "
   
   If Chk_FecAct.Value = 0 Then
      g_str_Parame = g_str_Parame & "     AND CAJMOV_FECDEP = '" & l_str_FecCar & "' "
   Else
      g_str_Parame = g_str_Parame & "     AND CAJMOV_FECDEP <= '" & l_str_FecCar & "' "
   End If
   g_str_Parame = g_str_Parame & "        AND CAJMOV_FLGPRO = 0 "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_TIPMOV = '1102' "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_TIPDOC IN (1, 6) "                      '1-DNI y 6-RUC
   g_str_Parame = g_str_Parame & "      ORDER BY A.CAJMOV_FECMOV , A.CAJMOV_NUMMOV "
   g_str_Parame = g_str_Parame & "      )"
   
   g_str_Parame = g_str_Parame & "      GROUP BY CAMPO_IDE_01 , CAMPO_IDE_02  , CAMPO_IDE_03 , CAMPO_IDE_04 , CAMPO_IDE_05   , CAMPO_IDE_06   , CAMPO_IDE_07   , CAMPO_IDE_08  , CAMPO_EMI_01  , CAMPO_EMI_02,"
   g_str_Parame = g_str_Parame & "               CAMPO_EMI_03 , CAMPO_EMI_04  , CAMPO_EMI_05 , CAMPO_EMI_06 , CAMPO_EMI_07   , CAMPO_EMI_08   , CAMPO_EMI_09   , CAMPO_EMI_10  , CAMPO_EMI_11  , CAMPO_EMI_12,"
   g_str_Parame = g_str_Parame & "               CAMPO_EMI_13 , CAMPO_EMI_14  , CAMPO_EMI_15 , CAMPO_REC_01 , CAMPO_REC_02   , CAMPO_REC_03   , CAMPO_REC_04   , CAMPO_REC_05  , CAMPO_REC_06  , CAMPO_REC_07,"
   g_str_Parame = g_str_Parame & "               CAMPO_REC_08 , CAMPO_REC_09  , CAMPO_REC_10 , CAMPO_REC_11 , CAMPO_REC_12   , CAMPO_DRF_01   , CAMPO_DRF_02   , CAMPO_DRF_03  , CAMPO_DRF_04  , CAMPO_DRF_05,"
   g_str_Parame = g_str_Parame & "               CAMPO_DRF_06 , CAMPO_CAB_01  , CAMPO_CAB_02 , CAMPO_CAB_03 , CAMPO_CAB_04   , CAMPO_CAB_05   , CAMPO_CAB_06   , CAMPO_CAB_07  , CAMPO_CAB_08  , CAMPO_CAB_09,"
   g_str_Parame = g_str_Parame & "               CAMPO_CAB_10 , CAMPO_CAB_11  , CAMPO_CAB_12 , CAMPO_CAB_13 , CAMPO_CAB_14   , CAMPO_CAB_15   , CAMPO_CAB_16   , CAMPO_CAB_17  , CAMPO_CAB_18_1, CAMPO_CAB_18_2,"
   g_str_Parame = g_str_Parame & "               CAMPO_CAB_19 , CAMPO_CAB_20  , CAMPO_CAB_21 , CAMPO_CAB_22 , CAMPO_CAB_23   , CAMPO_CAB_24   , CAMPO_CAB_25   , CAMPO_CAB_26  , CAMPO_CAB_27  , CAMPO_DET_01,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET_02 , CAMPO_DET_03  , CAMPO_DET_04 , "
   g_str_Parame = g_str_Parame & "               CAMPO_DET_06 , CAMPO_DET_07  , CAMPO_DET_08 , CAMPO_DET_09 , CAMPO_DET_10_1 , CAMPO_DET_10_2 , CAMPO_DET_10_3 , CAMPO_DET_11  , CAMPO_DET_12  , CAMPO_DET_13,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET_14 , CAMPO_DET_15  , CAMPO_DET_16 , CAMPO_DET_17 , CAMPO_DET_18   , CAMPO_DET_19   , CAMPO_DET_20   , CAMPO_DET_21  , CAMPO_DET_22  , CAMPO_DET_23,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET_24 , CAMPO_DET_25  , CAMPO_DET_26 , CAMPO_DET_27 , CAMPO_DET_28   , CAMPO_DET_29   , CAMPO_DET_30   , CAMPO_DET_31  , CAMPO_DET_32  , CAMPO_DET_33,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET_34 , CAMPO_DET_35  , CAMPO_DET_36 , CAMPO_DET_37 , CAMPO_DET2_01  , CAMPO_DET2_02  , CAMPO_DET2_03  , CAMPO_DET2_04 , "
   g_str_Parame = g_str_Parame & "               CAMPO_DET2_06, CAMPO_DET2_07, CAMPO_DET2_08 , CAMPO_DET2_09, CAMPO_DET2_10_1, CAMPO_DET2_10_2, CAMPO_DET2_10_3, CAMPO_DET2_11 , CAMPO_DET2_12 , CAMPO_DET2_13,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET2_14, CAMPO_DET2_15, CAMPO_DET2_16 , CAMPO_DET2_17, CAMPO_DET2_18  , CAMPO_DET2_19  , CAMPO_DET2_20  , CAMPO_DET2_21 , CAMPO_DET2_22 , CAMPO_DET2_23,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET2_24, CAMPO_DET2_25, CAMPO_DET2_26 , CAMPO_DET2_27, CAMPO_DET2_28  , CAMPO_DET2_29  , CAMPO_DET2_30  , CAMPO_DET2_31 , CAMPO_DET2_32 , CAMPO_DET2_33,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET2_34, CAMPO_DET2_35, CAMPO_DET2_36 , CAMPO_DET2_37, CAMPO_ADI_01   , CAMPO_ADI_02   , CAMPO_ADI_03   , CAMPO_ADI_04  , OPERACION     , NUMERO_MOVIMIENTO,"
   g_str_Parame = g_str_Parame & "               FECHA_CANCELACION , SITUACION, FECHA_DEPOSITO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta principal, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontró ningún registro.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro anterior en OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
    
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
      moddat_g_str_NumOpe = g_rst_Princi!OPERACION
      moddat_g_str_Codigo = g_rst_Princi!NUMERO_MOVIMIENTO
      
      If g_rst_Princi!SITUACION <> 2 Then
         If g_rst_Princi!FECHA_DEPOSITO <= g_rst_Princi!FECHA_CANCELACION Then
            GoTo Ingresar
         End If
      Else
      
Ingresar:
         Call fs_Obtener_Codigo("01", r_lng_Contad, r_int_SerFac, r_lng_NumFac)
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELE (      "
         g_str_Parame = g_str_Parame & " DOCELE_CODIGO                 , "
         g_str_Parame = g_str_Parame & " DOCELE_NUMOPE                 , "
         g_str_Parame = g_str_Parame & " DOCELE_NUMMOV                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FECPRO                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FECAUT                 , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_FECEMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_HOREMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPMON             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_NUMORC             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_FECVCT             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_NOMCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_UBIGEO             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_URBANI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_TELEMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_COREMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_TELREC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_CORREC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_CODMOT             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_DESMOT             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_REGPER      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_BASIMP_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOPER             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOTOT_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIMP             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOIMP             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_OTRCAR             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_TOTDSC      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTDSC             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_IMPTOT_DOCUME      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_DSCGLO             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_INFPPG             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTANT             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TIPOPE             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_LEYEND             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_TITADI             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_VALADI             , "
         g_str_Parame = g_str_Parame & " DOCELE_FLGENV                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FLGRPT                 , "
         g_str_Parame = g_str_Parame & " DOCELE_SITUAC                 , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES ( "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & " , "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
         g_str_Parame = g_str_Parame & "" & moddat_g_str_Codigo & " , "
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "     , "
         g_str_Parame = g_str_Parame & " NULL, "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_05 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_06 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_05 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_06 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_09 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_10 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_11 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_12 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_13 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_14 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_15 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_05 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_REC_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "'" & Mid(Replace(g_rst_Princi!CAMPO_REC_06, "  ", " "), 1, 100) & "'                                    , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_09 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_10 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_11 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_12 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_04 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DRF_05) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DRF_05 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_06 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_03 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_04) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_04 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_06 & "                    , "
         End If
               
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_08 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_09 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_10) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_10 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_11 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_12 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_13 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_14 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_15) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_15 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_16) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_16 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_17) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_17 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_18_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_18_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_18_2 & "                  , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_19) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_19 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_20 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_21 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_22) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_22 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_23) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_23 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_24 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_25) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_25 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_26 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_27 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_04 & "'                     , "
         g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
         g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
         g_str_Parame = g_str_Parame & "" & 1 & "                                               , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                               
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELE, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
   
         
         ''INTERES COMPENSATORIO
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES (                        "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_04 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_06 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_08 & "                    , "
         End If
   
         If IsNull(g_rst_Princi!CAMPO_DET_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_09 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_10_2 & "                  , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_3 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_11 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_12 & "                    , "
         End If
      
         If IsNull(g_rst_Princi!CAMPO_DET_13) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_13 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_14 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_16 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_17 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_18 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_19 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_20 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_21 & "   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_22 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_23 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_24 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_25) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_25 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_26 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_27 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_28 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_29 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_30 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_31 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_32 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_33 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_34 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_35 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_36 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_37 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "'                           , "
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de INTERES COMPENSATORIO, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
         
                                      
         ''OTROS IMPORTES
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES ("
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_03 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_04 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_05 & "'                    , "
         If IsNull(g_rst_Princi!CAMPO_DET2_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_06 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_07 & "'                    , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_08 & "                   , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_09 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_1 & "'                  , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_10_2 & "                 , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_3 & "'                  , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_11 & "                   , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_12 & "                   , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_13) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_13 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_14 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_16 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_17 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_18 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_19 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_20 & "'                    , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_21 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_22 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_23 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_24 & "'                    , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_25) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_25 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_26 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_27 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_28 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_29 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_30 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_31 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_32 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_33 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_34 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_35 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_36 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_37 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de OTROS IMPORTES, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
                                                                                                   
   
         DoEvents: DoEvents: DoEvents
         
         
         'ACTUALIZA EL CAMPO CAJMOV_FLGPRO PARA IDENTIFICAR CUALES SE HAN PROCESADO Y YA SE ENCUENTRAN EN CNTBL_DOCELE
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "UPDATE OPE_CAJMOV SET CAJMOV_FLGPRO = 1 "
         g_str_Parame = g_str_Parame & " WHERE CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' "
         g_str_Parame = g_str_Parame & "   AND CAJMOV_NUMMOV = '" & moddat_g_str_Codigo & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar CAJMOV_FLGPRO de la tabla OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
      
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Genera_FactAnterior")

End Sub
Private Sub fs_Generar_Boletas()
Dim r_lng_Contad     As Long
Dim r_int_SerFac     As Integer
Dim r_lng_NumFac     As Long

   On Error GoTo MyError
   
   Screen.MousePointer = 11
   
   g_str_Parame = ""
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = "03" Then
   
      g_str_Parame = g_str_Parame & "   SELECT  CAMPO_IDE_01, CAMPO_IDE_02  , CAMPO_IDE_03 , CAMPO_IDE_04 , CAMPO_IDE_05   , CAMPO_IDE_06   , CAMPO_IDE_07   , CAMPO_IDE_08 , CAMPO_EMI_01  , CAMPO_EMI_02,"
      g_str_Parame = g_str_Parame & "           CAMPO_EMI_03, CAMPO_EMI_04  , CAMPO_EMI_05 , CAMPO_EMI_06 , CAMPO_EMI_07   , CAMPO_EMI_08   , CAMPO_EMI_09   , CAMPO_EMI_10 , CAMPO_EMI_11  , CAMPO_EMI_12,"
      g_str_Parame = g_str_Parame & "           CAMPO_EMI_13, CAMPO_EMI_14  , CAMPO_EMI_15 , CAMPO_REC_01 , CAMPO_REC_02   , CAMPO_REC_03   , CAMPO_REC_04   , CAMPO_REC_05 , CAMPO_REC_06  , CAMPO_REC_07,"
      g_str_Parame = g_str_Parame & "           CAMPO_REC_08, CAMPO_REC_09  , CAMPO_REC_10 , CAMPO_REC_11 , CAMPO_REC_12   , CAMPO_DRF_01   , CAMPO_DRF_02   , CAMPO_DRF_03 , CAMPO_DRF_04  , CAMPO_DRF_05,"
      g_str_Parame = g_str_Parame & "           CAMPO_DRF_06, CAMPO_CAB_01  , CAMPO_CAB_02 , CAMPO_CAB_03 , CAMPO_CAB_04   , CAMPO_CAB_05   , CAMPO_CAB_06   , CAMPO_CAB_07 , CAMPO_CAB_08  , CAMPO_CAB_09,"
      g_str_Parame = g_str_Parame & "           CAMPO_CAB_10, CAMPO_CAB_11  , CAMPO_CAB_12 , CAMPO_CAB_13 , CAMPO_CAB_14   , CAMPO_CAB_15   , CAMPO_CAB_16   , CAMPO_CAB_17 , CAMPO_CAB_18_1, CAMPO_CAB_18_2,"
      g_str_Parame = g_str_Parame & "           CAMPO_CAB_19, CAMPO_CAB_20  , CAMPO_CAB_21 , CAMPO_CAB_22 , CAMPO_CAB_23   , CAMPO_CAB_24   , CAMPO_CAB_25   , CAMPO_CAB_26 , CAMPO_CAB_27  , CAMPO_DET_01,"
      g_str_Parame = g_str_Parame & "           CAMPO_DET_02, CAMPO_DET_03  , CAMPO_DET_04 , "
      g_str_Parame = g_str_Parame & "           MIN(CAMPO_DET_05) AS CAMPO_DET_05, "
      g_str_Parame = g_str_Parame & "           CAMPO_DET_06, CAMPO_DET_07  , CAMPO_DET_08 , CAMPO_DET_09 , CAMPO_DET_10_1 , CAMPO_DET_10_2 , CAMPO_DET_10_3 , CAMPO_DET_11 , CAMPO_DET_12  , CAMPO_DET_13,"
      g_str_Parame = g_str_Parame & "           CAMPO_DET_14, CAMPO_DET_15  , CAMPO_DET_16 , CAMPO_DET_17 , CAMPO_DET_18   , CAMPO_DET_19   , CAMPO_DET_20   , CAMPO_DET_21 , CAMPO_DET_22  , CAMPO_DET_23,"
      g_str_Parame = g_str_Parame & "           CAMPO_DET_24, CAMPO_DET_25  , CAMPO_DET_26 , CAMPO_DET_27 , CAMPO_DET_28   , CAMPO_DET_29   , CAMPO_DET_30   , CAMPO_DET_31 , CAMPO_DET_32  , CAMPO_DET_33,"
      g_str_Parame = g_str_Parame & "           CAMPO_DET_34, CAMPO_DET_35  , CAMPO_DET_36 , CAMPO_DET_37 , CAMPO_DET2_01  , CAMPO_DET2_02  , CAMPO_DET2_03  , CAMPO_DET2_04,"
      g_str_Parame = g_str_Parame & "           MIN(CAMPO_DET2_05) AS CAMPO_DET2_05, "
      g_str_Parame = g_str_Parame & "           CAMPO_DET2_06, CAMPO_DET2_07, CAMPO_DET2_08, CAMPO_DET2_09, CAMPO_DET2_10_1, CAMPO_DET2_10_2, CAMPO_DET2_10_3, CAMPO_DET2_11, CAMPO_DET2_12 , CAMPO_DET2_13,"
      g_str_Parame = g_str_Parame & "           CAMPO_DET2_14, CAMPO_DET2_15, CAMPO_DET2_16, CAMPO_DET2_17, CAMPO_DET2_18  , CAMPO_DET2_19  , CAMPO_DET2_20  , CAMPO_DET2_21, CAMPO_DET2_22 , CAMPO_DET2_23,"
      g_str_Parame = g_str_Parame & "           CAMPO_DET2_24, CAMPO_DET2_25, CAMPO_DET2_26, CAMPO_DET2_27, CAMPO_DET2_28  , CAMPO_DET2_29  , CAMPO_DET2_30  , CAMPO_DET2_31, CAMPO_DET2_32 , CAMPO_DET2_33,"
      g_str_Parame = g_str_Parame & "           CAMPO_DET2_34, CAMPO_DET2_35, CAMPO_DET2_36, CAMPO_DET2_37, CAMPO_ADI_01   , CAMPO_ADI_02   , CAMPO_ADI_03   , CAMPO_ADI_04 , OPERACION     , NUMERO_MOVIMIENTO,"
      g_str_Parame = g_str_Parame & "           FECHA_CANCELACION , SITUACION, FECHA_DEPOSITO "
      g_str_Parame = g_str_Parame & "     FROM ( "
   '''
      g_str_Parame = g_str_Parame & "     SELECT 'IDE'                                                                                               AS CAMPO_IDE_01, "
      g_str_Parame = g_str_Parame & "            'B'                                                                                                 AS CAMPO_IDE_02, "  '001-
      g_str_Parame = g_str_Parame & "            SUBSTR(CAJMOV_FECDEP,1,4) || '-' || SUBSTR(CAJMOV_FECDEP,5,2) || '-' || SUBSTR(CAJMOV_FECDEP,7,2)   AS CAMPO_IDE_03, "
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_04,"
      g_str_Parame = g_str_Parame & "            '03'                                                                                                AS CAMPO_IDE_05,              " '--CATALOGO N°1.- FACTURA
      g_str_Parame = g_str_Parame & "            C.CATSUN_CODIGO                                                                                     AS CAMPO_IDE_06,              " '--CATALOGO N°2.-
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_07,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_08,"
      g_str_Parame = g_str_Parame & "            'EMI'                                                                                               AS CAMPO_EMI_01,"
      g_str_Parame = g_str_Parame & "            'B'                                                                                                 AS CAMPO_EMI_02,"  '001-
      g_str_Parame = g_str_Parame & "            '6'                                                                                                 AS CAMPO_EMI_03,              " '--CATALOGO N°6.-
      g_str_Parame = g_str_Parame & "            '20511904162'                                                                                       AS CAMPO_EMI_04,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_05,"
      g_str_Parame = g_str_Parame & "            'EDPYME MICASITA SA'                                                                                AS CAMPO_EMI_06,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_07,              " '--ATALOGO N°13.-UBIGEO
      g_str_Parame = g_str_Parame & "            'AV RIVERA NAVARRETE 645'                                                                           AS CAMPO_EMI_08,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_09,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_10,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_11,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_12,"
      g_str_Parame = g_str_Parame & "            'PE'                                                                                                AS CAMPO_EMI_13,              " '--CATALOGO N°4.-
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_14,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_15,"
      g_str_Parame = g_str_Parame & "            'REC'                                                                                               AS CAMPO_REC_01,"
      g_str_Parame = g_str_Parame & "            'B'                                                                                                 AS CAMPO_REC_02,"  '001-
      g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_TIPDOC)                                                                               AS CAMPO_REC_03,              " '--CATALOGO N°6.-
      g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_NUMDOC)                                                                               AS CAMPO_REC_04,"
      g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_APEPAT) || ' ' || TRIM(D.DATGEN_APEMAT) || ' ' || TRIM(D.DATGEN_NOMBRE)               AS CAMPO_REC_05,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_REC_06,"
      g_str_Parame = g_str_Parame & "            TRIM(H.PARDES_DESCRI)                                                                               AS CAMPO_REC_07,"
      g_str_Parame = g_str_Parame & "            TRIM(I.PARDES_DESCRI)                                                                               AS CAMPO_REC_08,"
      g_str_Parame = g_str_Parame & "            TRIM(G.PARDES_DESCRI)                                                                               AS CAMPO_REC_09,"
      g_str_Parame = g_str_Parame & "            'PE'                                                                                                AS CAMPO_REC_10,"
      g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_TELEFO)                                                                               AS CAMPO_REC_11,"
      g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_DIRELE)                                                                               AS CAMPO_REC_12,"
      g_str_Parame = g_str_Parame & "            'DRF'                                                                                               AS CAMPO_DRF_01,"
      g_str_Parame = g_str_Parame & "            'B'                                                                                                 AS CAMPO_DRF_02," '001-
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_03,               "
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_04,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_05,               " '--PARA NOTA CREDITO/DEBITO
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_06,"
      g_str_Parame = g_str_Parame & "            'CAB'                                                                                               AS CAMPO_CAB_01,"
      g_str_Parame = g_str_Parame & "            'B'                                                                                                 AS CAMPO_CAB_02," '001-
      g_str_Parame = g_str_Parame & "            '1001'                                                                                              AS CAMPO_CAB_03,               " '--CATALOGO N°14.-
      g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_04,"
      g_str_Parame = g_str_Parame & "            '1002'                                                                                              AS CAMPO_CAB_05,               " '--CATALOGO N°14.-
      g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_06,"
      g_str_Parame = g_str_Parame & "            '1003'                                                                                              AS CAMPO_CAB_07,               " '--CATALOGO N°14.-
      g_str_Parame = g_str_Parame & "            '0.00'                                                                                              AS CAMPO_CAB_08,"
      g_str_Parame = g_str_Parame & "            '1004'                                                                                              AS CAMPO_CAB_09,               " '--CATALOGO N°14.-
      g_str_Parame = g_str_Parame & "            '0.00'                                                                                              AS CAMPO_CAB_10,"
      g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_CAB_11,               " '--CATALOGO N°14.-
      g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_12,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_13,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_14,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_15,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_16,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_17,"
      g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_CAB_18_1,"
      g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_18_2,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_19,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_20,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_21,"
      g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_22,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_23,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_24,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_25,"
      g_str_Parame = g_str_Parame & "            '13'                                                                                                AS CAMPO_CAB_26,                " '--CATALOGO N°17.-
      g_str_Parame = g_str_Parame & "            '[1000'                                                                                             AS CAMPO_CAB_27,                " '--CATALOGO N°15.- DETALLE EN LETRAS DEL IMPORTE
      g_str_Parame = g_str_Parame & "            'DET1'                                                                                              AS CAMPO_DET_01,"
      g_str_Parame = g_str_Parame & "            'B'                                                                                                 AS CAMPO_DET_02," '001-
      g_str_Parame = g_str_Parame & "            '001'                                                                                               AS CAMPO_DET_03,                " '-- Número de orden de ítem
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_04,                "
      
      g_str_Parame = g_str_Parame & "            'INTERES ' ||"
      g_str_Parame = g_str_Parame & "            SUBSTR(TO_CHAR(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(HIPCUO_FECVCT,1,4)  AS CAMPO_DET_05," 'CAJMOV_FECDEP
      g_str_Parame = g_str_Parame & "            1.000                                                                                               AS CAMPO_DET_06,"
      g_str_Parame = g_str_Parame & "            'NIU'                                                                                               AS CAMPO_DET_07,                " '--CATALOGO N°3.-
      g_str_Parame = g_str_Parame & "            A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_08,"
      g_str_Parame = g_str_Parame & "            1.000 * A.CAJMOV_INTERE                                                                             AS CAMPO_DET_09,"
      g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_DET_10_1,              " '--CATALOGO N°5,7 u 8.-
      g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_DET_10_2,"
      g_str_Parame = g_str_Parame & "            '30'                                                                                                AS CAMPO_DET_10_3,"
      g_str_Parame = g_str_Parame & "            A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_11,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_12,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_13,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_14,"
      g_str_Parame = g_str_Parame & "            '84121901'                                                                                          AS CAMPO_DET_15,                " '--CATALOGO N°15.-
      g_str_Parame = g_str_Parame & "            '7004'                                                                                              AS CAMPO_DET_16,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_17,                " '--NRO CONTRATO O PRESTAMO
      g_str_Parame = g_str_Parame & "            '7005'                                                                                              AS CAMPO_DET_18,"
      g_str_Parame = g_str_Parame & "            SUBSTR(J.HIPMAE_FECACT,1,4) || '-' || SUBSTR(J.HIPMAE_FECACT,5,2) || '-' || SUBSTR(J.HIPMAE_FECACT,7,2)"
      g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET_19,"
      g_str_Parame = g_str_Parame & "            '7001'                                                                                              AS CAMPO_DET_20,"
      g_str_Parame = g_str_Parame & "            '1'                                                                                                 AS CAMPO_DET_21,                " '--CATALOGO N°26.- Si es construcción /adquisición
      g_str_Parame = g_str_Parame & "            '7003'                                                                                              AS CAMPO_DET_22,"
      g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_23,"
      g_str_Parame = g_str_Parame & "            '7002'                                                                                              AS CAMPO_DET_24,"
      g_str_Parame = g_str_Parame & "            CASE WHEN L.SOLMAE_PRIVIV = 1 THEN 3  "
      g_str_Parame = g_str_Parame & "                 WHEN L.SOLMAE_PRIVIV = 2 THEN 0 END                                                            AS CAMPO_DET_25,                " '--CATALOGO N°27.- VERIFICAR
      g_str_Parame = g_str_Parame & "            '7007'                                                                                              AS CAMPO_DET_26,"
      g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_27,"
      g_str_Parame = g_str_Parame & "            '7006'                                                                                              AS CAMPO_DET_28,"
      g_str_Parame = g_str_Parame & "            '-'/*P.SOLINM_UBIGEO*/                                                                              AS CAMPO_DET_29,"
      g_str_Parame = g_str_Parame & "            '7008'                                                                                              AS CAMPO_DET_30,"
      g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_31,"
      g_str_Parame = g_str_Parame & "            '7011'                                                                                              AS CAMPO_DET_32,"
      g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_33,"
      g_str_Parame = g_str_Parame & "            '7009'                                                                                              AS CAMPO_DET_34,"
      g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_35,"
      g_str_Parame = g_str_Parame & "            '7010'                                                                                              AS CAMPO_DET_36,"
      g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_37,"
      g_str_Parame = g_str_Parame & "            'DET2'                                                                                              AS CAMPO_DET2_01,"
      g_str_Parame = g_str_Parame & "            'B'                                                                                                 AS CAMPO_DET2_02," '001-
      g_str_Parame = g_str_Parame & "            '002'                                                                                               AS CAMPO_DET2_03,                " '-- Número de orden de ítem
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_04,                "
      
      g_str_Parame = g_str_Parame & "            'OTROS IMPORTES ' ||"
      g_str_Parame = g_str_Parame & "            SUBSTR(TO_CHAR(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(HIPCUO_FECVCT,1,4)  AS CAMPO_DET2_05," 'CAJMOV_FECDEP
      g_str_Parame = g_str_Parame & "            1.000                                                                                               AS CAMPO_DET2_06,"
      g_str_Parame = g_str_Parame & "            'NIU'                                                                                               AS CAMPO_DET2_07,                " '--CATALOGO N°3.-
      g_str_Parame = g_str_Parame & "            (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)"
      g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_08,"
      g_str_Parame = g_str_Parame & "            1.000 * (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)"
      g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_09,"
      g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_DET2_10_1,              " '--CATALOGO N°5,7 u 8.-
      g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_DET2_10_2,"
      g_str_Parame = g_str_Parame & "            '30'                                                                                                AS CAMPO_DET2_10_3,"
      g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG - A.CAJMOV_INTERE "
      g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_11,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_12,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_13,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_14,"
      g_str_Parame = g_str_Parame & "            '84121501'                                                                                          AS CAMPO_DET2_15,                " '--CATALOGO N°15.-
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_16,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_17,                " '--NRO CONTRATO O PRESTAMO
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_18,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_19,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_20,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_21,                " '--CATALOGO N°26.- Si es construcción /adquisición
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_22,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_23,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_24,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_25,                " '--CATALOGO N°27.- VERIFICAR
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_26,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_27,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_28,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_29,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_30,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_31,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_32,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_33,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_34,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_35,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_36,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_37,"
   
      g_str_Parame = g_str_Parame & "            'ADI1'                                                                                              AS CAMPO_ADI_01,"
      g_str_Parame = g_str_Parame & "            'B'                                                                                                 AS CAMPO_ADI_02," '001-
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_ADI_03,"
      g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_ADI_04,"
      g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_NUMOPE)                                                                               AS OPERACION, "
      g_str_Parame = g_str_Parame & "            CAJMOV_NUMMOV                                                                                       AS NUMERO_MOVIMIENTO , "
      g_str_Parame = g_str_Parame & "            HIPMAE_FECCAN                                                                                       AS FECHA_CANCELACION ,"
      g_str_Parame = g_str_Parame & "            HIPMAE_SITUAC                                                                                       AS SITUACION, "
      g_str_Parame = g_str_Parame & "            CAJMOV_FECDEP                                                                                       AS FECHA_DEPOSITO "
      g_str_Parame = g_str_Parame & "       FROM OPE_CAJMOV A "
      g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '204' AND B.PARDES_CODITE = A.CAJMOV_MONPAG "
      g_str_Parame = g_str_Parame & "            INNER JOIN CNTBL_CATSUN C ON C.CATSUN_NROCAT = 2 AND C.CATSUN_DESCRI = TRIM(B.PARDES_DESCRI) "
      g_str_Parame = g_str_Parame & "            INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = A.CAJMOV_TIPDOC AND D.DATGEN_NUMDOC = A.CAJMOV_NUMDOC "
      g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 201 AND E.PARDES_CODITE = D.DATGEN_TIPVIA "
      g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 202 AND F.PARDES_CODITE = D.DATGEN_TIPZON "
      g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES G ON G.PARDES_CODGRP = 101 AND G.PARDES_CODITE = D.DATGEN_UBIGEO "
      g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 101 AND H.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,2)||'0000' "
      g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES I ON I.PARDES_CODGRP = 101 AND I.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "            INNER JOIN CRE_HIPMAE J ON J.HIPMAE_NUMOPE = A.CAJMOV_NUMOPE "
      g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_HIPGAR K ON K.HIPGAR_NUMOPE = J.HIPMAE_NUMOPE AND K.HIPGAR_BIEGAR = 1 "
      g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_SOLMAE L ON L.SOLMAE_NUMERO = J.HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_SOLINM P ON P.SOLINM_NUMSOL = J.HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES Q ON Q.PARDES_CODGRP = 201 AND Q.PARDES_CODITE = P.SOLINM_TIPVIA "
      g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES R ON R.PARDES_CODGRP = 202 AND R.PARDES_CODITE = P.SOLINM_TIPZON "
      g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES S ON S.PARDES_CODGRP = 101 AND S.PARDES_CODITE = P.SOLINM_UBIGEO "
      g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES T ON T.PARDES_CODGRP = 101 AND T.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES U ON U.PARDES_CODGRP = 101 AND U.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,2)||'0000' "
      
      g_str_Parame = g_str_Parame & "            INNER JOIN CRE_HIPPAG V ON V.HIPPAG_NUMOPE = A.CAJMOV_NUMOPE AND V.HIPPAG_FECPAG = A.CAJMOV_FECDEP AND V.HIPPAG_NUMMOV = A.CAJMOV_NUMMOV"
      g_str_Parame = g_str_Parame & "            INNER JOIN CRE_HIPCUO W ON W.HIPCUO_NUMOPE = V.HIPPAG_NUMOPE AND W.HIPCUO_TIPCRO = 1 AND W.HIPCUO_NUMCUO = V.HIPPAG_NUMCUO"
                  
      g_str_Parame = g_str_Parame & "      WHERE CAJMOV_SUCMOV IS NOT NULL "
      g_str_Parame = g_str_Parame & "        AND CAJMOV_USUMOV IS NOT NULL "
      g_str_Parame = g_str_Parame & "        AND CAJMOV_FECMOV > 0"
      g_str_Parame = g_str_Parame & "        AND CAJMOV_NUMMOV > 0 "
      g_str_Parame = g_str_Parame & "        AND CAJMOV_CODBAN IS NOT NULL "
      
      If Chk_FecAct.Value = 0 Then
         g_str_Parame = g_str_Parame & "     AND CAJMOV_FECDEP = '" & l_str_FecCar & "' "
      Else
         g_str_Parame = g_str_Parame & "     AND CAJMOV_FECDEP <= '" & l_str_FecCar & "' "
      End If
      g_str_Parame = g_str_Parame & "        AND CAJMOV_FLGPRO = 0 "
      g_str_Parame = g_str_Parame & "        AND CAJMOV_TIPMOV = '1102' "
      g_str_Parame = g_str_Parame & "        AND CAJMOV_TIPDOC IN (1, 6) "                      '1-DNI y 6-RUC
      g_str_Parame = g_str_Parame & "      ORDER BY A.CAJMOV_FECMOV , A.CAJMOV_NUMMOV "
      g_str_Parame = g_str_Parame & "      )"
      
      g_str_Parame = g_str_Parame & "      GROUP BY CAMPO_IDE_01 , CAMPO_IDE_02  , CAMPO_IDE_03 , CAMPO_IDE_04 , CAMPO_IDE_05   , CAMPO_IDE_06   , CAMPO_IDE_07   , CAMPO_IDE_08  , CAMPO_EMI_01  , CAMPO_EMI_02,"
      g_str_Parame = g_str_Parame & "               CAMPO_EMI_03 , CAMPO_EMI_04  , CAMPO_EMI_05 , CAMPO_EMI_06 , CAMPO_EMI_07   , CAMPO_EMI_08   , CAMPO_EMI_09   , CAMPO_EMI_10  , CAMPO_EMI_11  , CAMPO_EMI_12,"
      g_str_Parame = g_str_Parame & "               CAMPO_EMI_13 , CAMPO_EMI_14  , CAMPO_EMI_15 , CAMPO_REC_01 , CAMPO_REC_02   , CAMPO_REC_03   , CAMPO_REC_04   , CAMPO_REC_05  , CAMPO_REC_06  , CAMPO_REC_07,"
      g_str_Parame = g_str_Parame & "               CAMPO_REC_08 , CAMPO_REC_09  , CAMPO_REC_10 , CAMPO_REC_11 , CAMPO_REC_12   , CAMPO_DRF_01   , CAMPO_DRF_02   , CAMPO_DRF_03  , CAMPO_DRF_04  , CAMPO_DRF_05,"
      g_str_Parame = g_str_Parame & "               CAMPO_DRF_06 , CAMPO_CAB_01  , CAMPO_CAB_02 , CAMPO_CAB_03 , CAMPO_CAB_04   , CAMPO_CAB_05   , CAMPO_CAB_06   , CAMPO_CAB_07  , CAMPO_CAB_08  , CAMPO_CAB_09,"
      g_str_Parame = g_str_Parame & "               CAMPO_CAB_10 , CAMPO_CAB_11  , CAMPO_CAB_12 , CAMPO_CAB_13 , CAMPO_CAB_14   , CAMPO_CAB_15   , CAMPO_CAB_16   , CAMPO_CAB_17  , CAMPO_CAB_18_1, CAMPO_CAB_18_2,"
      g_str_Parame = g_str_Parame & "               CAMPO_CAB_19 , CAMPO_CAB_20  , CAMPO_CAB_21 , CAMPO_CAB_22 , CAMPO_CAB_23   , CAMPO_CAB_24   , CAMPO_CAB_25   , CAMPO_CAB_26  , CAMPO_CAB_27  , CAMPO_DET_01,"
      g_str_Parame = g_str_Parame & "               CAMPO_DET_02 , CAMPO_DET_03  , CAMPO_DET_04 , "
      g_str_Parame = g_str_Parame & "               CAMPO_DET_06 , CAMPO_DET_07  , CAMPO_DET_08 , CAMPO_DET_09 , CAMPO_DET_10_1 , CAMPO_DET_10_2 , CAMPO_DET_10_3 , CAMPO_DET_11  , CAMPO_DET_12  , CAMPO_DET_13,"
      g_str_Parame = g_str_Parame & "               CAMPO_DET_14 , CAMPO_DET_15  , CAMPO_DET_16 , CAMPO_DET_17 , CAMPO_DET_18   , CAMPO_DET_19   , CAMPO_DET_20   , CAMPO_DET_21  , CAMPO_DET_22  , CAMPO_DET_23,"
      g_str_Parame = g_str_Parame & "               CAMPO_DET_24 , CAMPO_DET_25  , CAMPO_DET_26 , CAMPO_DET_27 , CAMPO_DET_28   , CAMPO_DET_29   , CAMPO_DET_30   , CAMPO_DET_31  , CAMPO_DET_32  , CAMPO_DET_33,"
      g_str_Parame = g_str_Parame & "               CAMPO_DET_34 , CAMPO_DET_35  , CAMPO_DET_36 , CAMPO_DET_37 , CAMPO_DET2_01  , CAMPO_DET2_02  , CAMPO_DET2_03  , CAMPO_DET2_04 , "
      g_str_Parame = g_str_Parame & "               CAMPO_DET2_06, CAMPO_DET2_07 , CAMPO_DET2_08, CAMPO_DET2_09, CAMPO_DET2_10_1, CAMPO_DET2_10_2, CAMPO_DET2_10_3, CAMPO_DET2_11 , CAMPO_DET2_12 , CAMPO_DET2_13,"
      g_str_Parame = g_str_Parame & "               CAMPO_DET2_14, CAMPO_DET2_15 , CAMPO_DET2_16, CAMPO_DET2_17, CAMPO_DET2_18  , CAMPO_DET2_19  , CAMPO_DET2_20  , CAMPO_DET2_21 , CAMPO_DET2_22 , CAMPO_DET2_23,"
      g_str_Parame = g_str_Parame & "               CAMPO_DET2_24, CAMPO_DET2_25 , CAMPO_DET2_26, CAMPO_DET2_27, CAMPO_DET2_28  , CAMPO_DET2_29  , CAMPO_DET2_30  , CAMPO_DET2_31 , CAMPO_DET2_32 , CAMPO_DET2_33,"
      g_str_Parame = g_str_Parame & "               CAMPO_DET2_34, CAMPO_DET2_35 , CAMPO_DET2_36, CAMPO_DET2_37, CAMPO_ADI_01   , CAMPO_ADI_02   , CAMPO_ADI_03   , CAMPO_ADI_04  , OPERACION     , NUMERO_MOVIMIENTO,"
      g_str_Parame = g_str_Parame & "               FECHA_CANCELACION , SITUACION, FECHA_DEPOSITO"
   
   
   ElseIf cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = "05" Then
   
      g_str_Parame = g_str_Parame & "        SELECT 'IDE'                                                                                               AS CAMPO_IDE_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_IDE_02, "
      g_str_Parame = g_str_Parame & "               SUBSTR(PPGCAB_FECPPG,1,4) || '-' || SUBSTR(PPGCAB_FECPPG,5,2) || '-' || SUBSTR(PPGCAB_FECPPG,7,2)   AS CAMPO_IDE_03, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_IDE_04, "
      g_str_Parame = g_str_Parame & "               '03'                                                                                                AS CAMPO_IDE_05, "
      g_str_Parame = g_str_Parame & "               C.CATSUN_CODIGO                                                                                     AS CAMPO_IDE_06, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_IDE_07, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_IDE_08, "
      g_str_Parame = g_str_Parame & "               'EMI'                                                                                               AS CAMPO_EMI_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_EMI_02, "
      g_str_Parame = g_str_Parame & "               '6'                                                                                                 AS CAMPO_EMI_03, "
      g_str_Parame = g_str_Parame & "               '20511904162'                                                                                       AS CAMPO_EMI_04, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_05, "
      g_str_Parame = g_str_Parame & "               'EDPYME MICASITA SA'                                                                                AS CAMPO_EMI_06, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_07, "
      g_str_Parame = g_str_Parame & "               'AV RIVERA NAVARRETE 645'                                                                           AS CAMPO_EMI_08, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_09, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_10, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_11, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_12, "
      g_str_Parame = g_str_Parame & "               'PE'                                                                                                AS CAMPO_EMI_13, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_14, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_15, "
      g_str_Parame = g_str_Parame & "               'REC'                                                                                               AS CAMPO_REC_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_REC_02, "
      g_str_Parame = g_str_Parame & "               TRIM(J.HIPMAE_TDOCLI)                                                                               AS CAMPO_REC_03, "
      g_str_Parame = g_str_Parame & "               TRIM(J.HIPMAE_NDOCLI)                                                                               AS CAMPO_REC_04, "
      g_str_Parame = g_str_Parame & "               TRIM(D.DATGEN_APEPAT) || ' ' || TRIM(D.DATGEN_APEMAT) || ' ' || TRIM(D.DATGEN_NOMBRE)               AS CAMPO_REC_05, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_REC_06, "
      g_str_Parame = g_str_Parame & "               TRIM(H.PARDES_DESCRI)                                                                               AS CAMPO_REC_07, "
      g_str_Parame = g_str_Parame & "               TRIM(I.PARDES_DESCRI)                                                                               AS CAMPO_REC_08, "
      g_str_Parame = g_str_Parame & "               TRIM(G.PARDES_DESCRI)                                                                               AS CAMPO_REC_09, "
      g_str_Parame = g_str_Parame & "               'PE'                                                                                                AS CAMPO_REC_10, "
      g_str_Parame = g_str_Parame & "               TRIM(D.DATGEN_TELEFO)                                                                               AS CAMPO_REC_11, "
      g_str_Parame = g_str_Parame & "               TRIM(D.DATGEN_DIRELE)                                                                               AS CAMPO_REC_12, "
      g_str_Parame = g_str_Parame & "               'DRF'                                                                                               AS CAMPO_DRF_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_DRF_02, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DRF_03, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DRF_04, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DRF_05, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DRF_06, "
      g_str_Parame = g_str_Parame & "               'CAB'                                                                                               AS CAMPO_CAB_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_CAB_02, "
      g_str_Parame = g_str_Parame & "               '1001'                                                                                              AS CAMPO_CAB_03, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_CAB_04, "
      g_str_Parame = g_str_Parame & "               '1002'                                                                                              AS CAMPO_CAB_05, "
      g_str_Parame = g_str_Parame & "               CASE WHEN A.PPGCAB_TIPPPG = 1 THEN A.PPGCAB_MTODEP ELSE PPGCAB_MTOTOT END                           AS CAMPO_CAB_06, "
      g_str_Parame = g_str_Parame & "               '1003'                                                                                              AS CAMPO_CAB_07, "
      g_str_Parame = g_str_Parame & "               '0.00'                                                                                              AS CAMPO_CAB_08, "
      g_str_Parame = g_str_Parame & "               '1004'                                                                                              AS CAMPO_CAB_09, "
      g_str_Parame = g_str_Parame & "               '0.00'                                                                                              AS CAMPO_CAB_10, "
      g_str_Parame = g_str_Parame & "               '1000'                                                                                              AS CAMPO_CAB_11, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_CAB_12, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_13, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_14, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_15, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_16, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_17, "
      g_str_Parame = g_str_Parame & "               '1000'                                                                                              AS CAMPO_CAB_18_1, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_CAB_18_2, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_19, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_20, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_21, "
      g_str_Parame = g_str_Parame & "               CASE WHEN A.PPGCAB_TIPPPG = 1 THEN A.PPGCAB_MTODEP ELSE PPGCAB_MTOTOT END                           AS CAMPO_CAB_22, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_23, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_24, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_25, "
      g_str_Parame = g_str_Parame & "               '13'                                                                                                AS CAMPO_CAB_26, "
      g_str_Parame = g_str_Parame & "               '[1000'                                                                                             AS CAMPO_CAB_27, "
      g_str_Parame = g_str_Parame & "               'DET1'                                                                                              AS CAMPO_DET_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_DET_02, "
      g_str_Parame = g_str_Parame & "               '001'                                                                                               AS CAMPO_DET_03, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_04, "
      g_str_Parame = g_str_Parame & "               'INTERES ' ||"
      g_str_Parame = g_str_Parame & "               SUBSTR(TO_CHAR(TO_DATE(PPGCAB_FECPPG,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(PPGCAB_FECPPG,1,4)  AS CAMPO_DET_05, "
      g_str_Parame = g_str_Parame & "               1.000                                                                                               AS CAMPO_DET_06, "
      g_str_Parame = g_str_Parame & "               'NIU'                                                                                               AS CAMPO_DET_07, "
      g_str_Parame = g_str_Parame & "               (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)                                                          AS CAMPO_DET_08, "
      g_str_Parame = g_str_Parame & "               1.000 * (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)                                                  AS CAMPO_DET_09, "
      g_str_Parame = g_str_Parame & "               '1000'                                                                                              AS CAMPO_DET_10_1, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_DET_10_2, "
      g_str_Parame = g_str_Parame & "               '30'                                                                                                AS CAMPO_DET_10_3, "
      g_str_Parame = g_str_Parame & "               (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)                                                          AS CAMPO_DET_11, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_12, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_13, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_14, "
      g_str_Parame = g_str_Parame & "               '84121901'                                                                                          AS CAMPO_DET_15, "
      g_str_Parame = g_str_Parame & "               '7004'                                                                                              AS CAMPO_DET_16, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_17, "
      g_str_Parame = g_str_Parame & "               '7005'                                                                                              AS CAMPO_DET_18, "
      g_str_Parame = g_str_Parame & "               SUBSTR(J.HIPMAE_FECACT,1,4) || '-' || SUBSTR(J.HIPMAE_FECACT,5,2) || '-' || SUBSTR(J.HIPMAE_FECACT,7,2) "
      g_str_Parame = g_str_Parame & "                                                                                                                   AS CAMPO_DET_19, "
      g_str_Parame = g_str_Parame & "               '7001'                                                                                              AS CAMPO_DET_20, "
      g_str_Parame = g_str_Parame & "               '1'                                                                                                 AS CAMPO_DET_21, "
      g_str_Parame = g_str_Parame & "               '7003'                                                                                              AS CAMPO_DET_22, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_23, "
      g_str_Parame = g_str_Parame & "               '7002'                                                                                              AS CAMPO_DET_24, "
      g_str_Parame = g_str_Parame & "               CASE WHEN L.SOLMAE_PRIVIV = 1 THEN 3  "
      g_str_Parame = g_str_Parame & "                    WHEN L.SOLMAE_PRIVIV = 2 THEN 0 END                                                            AS CAMPO_DET_25, "
      g_str_Parame = g_str_Parame & "               '7007'                                                                                              AS CAMPO_DET_26, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_27, "
      g_str_Parame = g_str_Parame & "               '7006'                                                                                              AS CAMPO_DET_28, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_29, "
      g_str_Parame = g_str_Parame & "               '7008'                                                                                              AS CAMPO_DET_30, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_31, "
      g_str_Parame = g_str_Parame & "               '7011'                                                                                              AS CAMPO_DET_32, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_33, "
      g_str_Parame = g_str_Parame & "               '7009'                                                                                              AS CAMPO_DET_34, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_35, "
      g_str_Parame = g_str_Parame & "               '7010'                                                                                              AS CAMPO_DET_36, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_37, "
      g_str_Parame = g_str_Parame & "               'DET2'                                                                                              AS CAMPO_DET2_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_DET2_02, "
      g_str_Parame = g_str_Parame & "               '002'                                                                                               AS CAMPO_DET2_03, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_04, "
      g_str_Parame = g_str_Parame & "               'OTROS IMPORTES ' || "
      g_str_Parame = g_str_Parame & "               SUBSTR(TO_CHAR(TO_DATE(A.PPGCAB_FECPPG,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(PPGCAB_FECPPG,1,4)  AS CAMPO_DET2_05, "
      g_str_Parame = g_str_Parame & "               1.000                                                                                               AS CAMPO_DET2_06, "
      g_str_Parame = g_str_Parame & "               'NIU'                                                                                               AS CAMPO_DET2_07, "
      g_str_Parame = g_str_Parame & "               ((CASE WHEN A.PPGCAB_TIPPPG = 1 THEN A.PPGCAB_MTODEP ELSE PPGCAB_MTOTOT END) - (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)) "
      g_str_Parame = g_str_Parame & "                                                                                                                   AS CAMPO_DET2_08, "
      g_str_Parame = g_str_Parame & "               1.000 * ((CASE WHEN A.PPGCAB_TIPPPG = 1 THEN A.PPGCAB_MTODEP ELSE PPGCAB_MTOTOT END) - (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)) "
      g_str_Parame = g_str_Parame & "                                                                                                                   AS CAMPO_DET2_09, "
      g_str_Parame = g_str_Parame & "               '1000'                                                                                              AS CAMPO_DET2_10_1, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_DET2_10_2, "
      g_str_Parame = g_str_Parame & "               '30'                                                                                                AS CAMPO_DET2_10_3, "
      g_str_Parame = g_str_Parame & "               (CASE WHEN A.PPGCAB_TIPPPG = 1 THEN A.PPGCAB_MTODEP ELSE PPGCAB_MTOTOT END) - (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC) "
      g_str_Parame = g_str_Parame & "                                                                                                                   AS CAMPO_DET2_11, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_12, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_13, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_14, "
      g_str_Parame = g_str_Parame & "               '84121501'                                                                                          AS CAMPO_DET2_15, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_16, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_17, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_18, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_19, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_20, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_21, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_22, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_23, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_24, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_25, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_26, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_27, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_28, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_29, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_30, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_31, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_32, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_33, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_34, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_35, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_36, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_37, "
      g_str_Parame = g_str_Parame & "               'ADI1'                                                                                              AS CAMPO_ADI_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_ADI_02, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_ADI_03, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_ADI_04, "
      g_str_Parame = g_str_Parame & "               TRIM(A.PPGCAB_NUMOPE)                                                                               AS OPERACION, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS NUMERO_MOVIMIENTO, "
      g_str_Parame = g_str_Parame & "               HIPMAE_FECCAN                                                                                       AS FECHA_CANCELACION, "
      g_str_Parame = g_str_Parame & "               HIPMAE_SITUAC                                                                                       AS SITUACION, "
      g_str_Parame = g_str_Parame & "               PPGCAB_FECPPG                                                                                       AS FECHA_DEPOSITO "
      g_str_Parame = g_str_Parame & "          FROM CRE_PPGCAB A "
      g_str_Parame = g_str_Parame & "               INNER JOIN CRE_HIPMAE J ON J.HIPMAE_NUMOPE = A.PPGCAB_NUMOPE "
      g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '204' AND B.PARDES_CODITE = J.HIPMAE_MONEDA "
      g_str_Parame = g_str_Parame & "               INNER JOIN CNTBL_CATSUN C ON C.CATSUN_NROCAT = 2 AND C.CATSUN_DESCRI = TRIM(B.PARDES_DESCRI) "
      g_str_Parame = g_str_Parame & "               INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = J.HIPMAE_TDOCLI AND D.DATGEN_NUMDOC = J.HIPMAE_NDOCLI"
      g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 201 AND E.PARDES_CODITE = D.DATGEN_TIPVIA "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 202 AND F.PARDES_CODITE = D.DATGEN_TIPZON "
      g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES G ON G.PARDES_CODGRP = 101 AND G.PARDES_CODITE = D.DATGEN_UBIGEO "
      g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 101 AND H.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,2)||'0000' "
      g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES I ON I.PARDES_CODGRP = 101 AND I.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "                LEFT JOIN CRE_SOLMAE L ON L.SOLMAE_NUMERO = J.HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & "                LEFT JOIN CRE_SOLINM P ON P.SOLINM_NUMSOL = J.HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES Q ON Q.PARDES_CODGRP = 201 AND Q.PARDES_CODITE = P.SOLINM_TIPVIA "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES R ON R.PARDES_CODGRP = 202 AND R.PARDES_CODITE = P.SOLINM_TIPZON "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES S ON S.PARDES_CODGRP = 101 AND S.PARDES_CODITE = P.SOLINM_UBIGEO "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES T ON T.PARDES_CODGRP = 101 AND T.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES U ON U.PARDES_CODGRP = 101 AND U.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,2)||'0000' "
      g_str_Parame = g_str_Parame & "         WHERE "
      g_str_Parame = g_str_Parame & "               A.PPGCAB_FECPPG >= 20180801 "
      g_str_Parame = g_str_Parame & "           AND A.PPGCAB_FECPPG <= 20190331 "
      g_str_Parame = g_str_Parame & "           AND A.PPGCAB_FLGPRO = 0 "
      g_str_Parame = g_str_Parame & "           AND HIPMAE_TDOCLI IN (1, 6) "
      'g_str_Parame = g_str_Parame & "           AND HIPMAE_NUMOPE IN ('0071300013') "
      g_str_Parame = g_str_Parame & "         ORDER BY A.PPGCAB_FECPPG "

   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta principal, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontró ningún registro.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro anterior en OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
    
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
      moddat_g_str_NumOpe = g_rst_Princi!OPERACION
      If Not IsNull(g_rst_Princi!NUMERO_MOVIMIENTO) Then
         moddat_g_str_Codigo = g_rst_Princi!NUMERO_MOVIMIENTO
      Else
         moddat_g_str_Codigo = 0
      End If
      
      If g_rst_Princi!SITUACION <> 2 Then
         If g_rst_Princi!FECHA_DEPOSITO <= g_rst_Princi!FECHA_CANCELACION Then
            GoTo Ingresar
         End If
      Else
      
Ingresar:
         Call fs_Obtener_Codigo("03", r_lng_Contad, r_int_SerFac, r_lng_NumFac)
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELE (      "
         g_str_Parame = g_str_Parame & " DOCELE_CODIGO                 , "
         g_str_Parame = g_str_Parame & " DOCELE_NUMOPE                 , "
         g_str_Parame = g_str_Parame & " DOCELE_NUMMOV                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FECPRO                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FECAUT                 , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_FECEMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_HOREMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPMON             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_NUMORC             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_FECVCT             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_NOMCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_UBIGEO             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_URBANI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_TELEMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_COREMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_TELREC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_CORREC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_CODMOT             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_DESMOT             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_REGPER      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_BASIMP_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOPER             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOTOT_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIMP             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOIMP             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_OTRCAR             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_TOTDSC      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTDSC             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_IMPTOT_DOCUME      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_DSCGLO             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_INFPPG             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTANT             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TIPOPE             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_LEYEND             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_TITADI             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_VALADI             , "
         g_str_Parame = g_str_Parame & " DOCELE_FLGENV                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FLGRPT                 , "
         g_str_Parame = g_str_Parame & " DOCELE_SITUAC                 , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES ( "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & " , "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
         g_str_Parame = g_str_Parame & "" & moddat_g_str_Codigo & " , "
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "     , "
         g_str_Parame = g_str_Parame & " NULL, "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_05 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_06 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_05 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_06 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_09 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_10 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_11 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_12 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_13 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_14 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_15 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_05 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_REC_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "'" & Mid(Replace(g_rst_Princi!CAMPO_REC_06, "  ", " "), 1, 100) & "'                                    , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_09 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_10 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_11 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_12 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_04 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DRF_05) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DRF_05 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_06 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_03 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_04) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_04 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_06 & "                    , "
         End If
               
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_08 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_09 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_10) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_10 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_11 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_12 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_13 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_14 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_15) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_15 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_16) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_16 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_17) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_17 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_18_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_18_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_18_2 & "                  , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_19) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_19 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_20 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_21 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_22) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_22 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_23) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_23 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_24 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_25) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_25 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_26 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_27 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_04 & "'                     , "
         g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
         g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
         g_str_Parame = g_str_Parame & "" & 1 & "                                               , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                               
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELE, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
   
         
         ''INTERES COMPENSATORIO
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES (                        "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_04 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_06 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_08 & "                    , "
         End If
   
         If IsNull(g_rst_Princi!CAMPO_DET_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_09 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_10_2 & "                  , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_3 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_11 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_12 & "                    , "
         End If
      
         If IsNull(g_rst_Princi!CAMPO_DET_13) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_13 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_14 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_16 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_17 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_18 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_19 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_20 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_21 & "   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_22 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_23 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_24 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_25) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_25 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_26 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_27 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_28 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_29 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_30 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_31 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_32 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_33 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_34 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_35 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_36 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_37 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "'                           , "
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de INTERES COMPENSATORIO, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
         
                                      
         ''OTROS IMPORTES
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES ("
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_03 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_04 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_05 & "'                    , "
         If IsNull(g_rst_Princi!CAMPO_DET2_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_06 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_07 & "'                    , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_08 & "                   , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_09 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_1 & "'                  , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_10_2 & "                 , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_3 & "'                  , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_11 & "                   , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_12 & "                   , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_13) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_13 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_14 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_16 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_17 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_18 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_19 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_20 & "'                    , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_21 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_22 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_23 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_24 & "'                    , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_25) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_25 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_26 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_27 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_28 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_29 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_30 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_31 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_32 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_33 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_34 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_35 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_36 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_37 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de OTROS IMPORTES, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
                                                                                                   
   
         DoEvents: DoEvents: DoEvents
         
         
         'ACTUALIZA EL CAMPO CAJMOV_FLGPRO PARA IDENTIFICAR CUALES SE HAN PROCESADO Y YA SE ENCUENTRAN EN CNTBL_DOCELE
         g_str_Parame = ""
         
         If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = "03" Then
            g_str_Parame = g_str_Parame & "UPDATE OPE_CAJMOV SET CAJMOV_FLGPRO = 1 "
            g_str_Parame = g_str_Parame & " WHERE CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            g_str_Parame = g_str_Parame & "   AND CAJMOV_NUMMOV = '" & moddat_g_str_Codigo & "' "
         Else
            g_str_Parame = g_str_Parame & "UPDATE CRE_PPGCAB SET PPGCAB_FLGPRO = 1 "
            g_str_Parame = g_str_Parame & " WHERE PPGCAB_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            g_str_Parame = g_str_Parame & "   AND PPGCAB_FECPPG = '" & g_rst_Princi!FECHA_DEPOSITO & "' " 'moddat_g_str_Codigo
         End If
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar CAJMOV_FLGPRO de la tabla OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
      
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Genera_FactAnterior")

End Sub
Private Sub fs_Generar_Facturas_NUEVO_FORMATO()
Dim r_lng_Contad     As Long
Dim r_int_SerFac     As Integer
Dim r_lng_NumFac     As Long

   On Error GoTo MyError
   
   Screen.MousePointer = 11
   
   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "   SELECT  CAMPO_IDE_01, CAMPO_IDE_02  , CAMPO_IDE_03 , CAMPO_IDE_04 , CAMPO_IDE_05   , CAMPO_IDE_06   , CAMPO_IDE_07   , CAMPO_IDE_08 , CAMPO_EMI_01  , CAMPO_EMI_02,"
'   g_str_Parame = g_str_Parame & "           CAMPO_EMI_03, CAMPO_EMI_04  , CAMPO_EMI_05 , CAMPO_EMI_06 , CAMPO_EMI_07   , CAMPO_EMI_08   , CAMPO_EMI_09   , CAMPO_EMI_10 , CAMPO_EMI_11  , CAMPO_EMI_12,"
'   g_str_Parame = g_str_Parame & "           CAMPO_EMI_13, CAMPO_EMI_14  , CAMPO_EMI_15 , CAMPO_REC_01 , CAMPO_REC_02   , CAMPO_REC_03   , CAMPO_REC_04   , CAMPO_REC_05 , CAMPO_REC_06  , CAMPO_REC_07,"
'   g_str_Parame = g_str_Parame & "           CAMPO_REC_08, CAMPO_REC_09  , CAMPO_REC_10 , CAMPO_REC_11 , CAMPO_REC_12   , CAMPO_DRF_01   , CAMPO_DRF_02   , CAMPO_DRF_03 , CAMPO_DRF_04  , CAMPO_DRF_05,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DRF_06, CAMPO_CAB_01  , CAMPO_CAB_02 , CAMPO_CAB_03 , CAMPO_CAB_04   , CAMPO_CAB_05   , CAMPO_CAB_06   , CAMPO_CAB_07 , CAMPO_CAB_08  , CAMPO_CAB_09,"
'   g_str_Parame = g_str_Parame & "           CAMPO_CAB_10, CAMPO_CAB_11  , CAMPO_CAB_12 , CAMPO_CAB_13 , CAMPO_CAB_14   , CAMPO_CAB_15   , CAMPO_CAB_16   , CAMPO_CAB_17 , CAMPO_CAB_18_1, CAMPO_CAB_18_2,"
'   g_str_Parame = g_str_Parame & "           CAMPO_CAB_19, CAMPO_CAB_20  , CAMPO_CAB_21 , CAMPO_CAB_22 , CAMPO_CAB_23   , CAMPO_CAB_24   , CAMPO_CAB_25   , CAMPO_CAB_26 , CAMPO_CAB_27  , CAMPO_DET_01,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET_02, CAMPO_DET_03  , CAMPO_DET_04 , "
'   g_str_Parame = g_str_Parame & "           MIN(CAMPO_DET_05) AS CAMPO_DET_05, "
'   g_str_Parame = g_str_Parame & "           CAMPO_DET_06, CAMPO_DET_07  , CAMPO_DET_08 , CAMPO_DET_09 , CAMPO_DET_10_1 , CAMPO_DET_10_2 , CAMPO_DET_10_3 , CAMPO_DET_11 , CAMPO_DET_12  , CAMPO_DET_13,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET_14, CAMPO_DET_15  , CAMPO_DET_16 , CAMPO_DET_17 , CAMPO_DET_18   , CAMPO_DET_19   , CAMPO_DET_20   , CAMPO_DET_21 , CAMPO_DET_22  , CAMPO_DET_23,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET_24, CAMPO_DET_25  , CAMPO_DET_26 , CAMPO_DET_27 , CAMPO_DET_28   , CAMPO_DET_29   , CAMPO_DET_30   , CAMPO_DET_31 , CAMPO_DET_32  , CAMPO_DET_33,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET_34, CAMPO_DET_35  , CAMPO_DET_36 , CAMPO_DET_37 , CAMPO_DET2_01  , CAMPO_DET2_02  , CAMPO_DET2_03  , CAMPO_DET2_04,"
'   g_str_Parame = g_str_Parame & "           MIN(CAMPO_DET2_05) AS CAMPO_DET2_05, "
'   g_str_Parame = g_str_Parame & "           CAMPO_DET2_06, CAMPO_DET2_07, CAMPO_DET2_08, CAMPO_DET2_09, CAMPO_DET2_10_1, CAMPO_DET2_10_2, CAMPO_DET2_10_3, CAMPO_DET2_11, CAMPO_DET2_12 , CAMPO_DET2_13,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET2_14, CAMPO_DET2_15, CAMPO_DET2_16, CAMPO_DET2_17, CAMPO_DET2_18  , CAMPO_DET2_19  , CAMPO_DET2_20  , CAMPO_DET2_21, CAMPO_DET2_22 , CAMPO_DET2_23,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET2_24, CAMPO_DET2_25, CAMPO_DET2_26, CAMPO_DET2_27, CAMPO_DET2_28  , CAMPO_DET2_29  , CAMPO_DET2_30  , CAMPO_DET2_31, CAMPO_DET2_32 , CAMPO_DET2_33,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET2_34, CAMPO_DET2_35, CAMPO_DET2_36, CAMPO_DET2_37, CAMPO_ADI_01   , CAMPO_ADI_02   , CAMPO_ADI_03   , CAMPO_ADI_04 , OPERACION     , NUMERO_MOVIMIENTO,"
'   g_str_Parame = g_str_Parame & "           FECHA_CANCELACION , SITUACION, FECHA_DEPOSITO "

   g_str_Parame = g_str_Parame & "    SELECT CAMPO_IDE_01,      CAMPO_IDE_02,       CAMPO_IDE_03,       CAMPO_IDE_04,       CAMPO_IDE_05,       CAMPO_IDE_06,      CAMPO_IDE_07,      CAMPO_IDE_08,      CAMPO_EMI_01,      CAMPO_EMI_02,"
   g_str_Parame = g_str_Parame & "           CAMPO_EMI_03,      CAMPO_EMI_04,       CAMPO_EMI_05,       CAMPO_EMI_06,       CAMPO_EMI_07,       CAMPO_EMI_08,      CAMPO_EMI_09,      CAMPO_EMI_10,      CAMPO_EMI_11,      CAMPO_EMI_12,"
   g_str_Parame = g_str_Parame & "           CAMPO_EMI_13,      CAMPO_EMI_14,       CAMPO_EMI_15,       CAMPO_EMI_16,       CAMPO_REC_01,       CAMPO_REC_02,      CAMPO_REC_03,      CAMPO_REC_04,      CAMPO_REC_05,      CAMPO_REC_06,"
   g_str_Parame = g_str_Parame & "           CAMPO_REC_07,      CAMPO_REC_08,       CAMPO_REC_09,       CAMPO_REC_10,       CAMPO_REC_11,       CAMPO_REC_12,      CAMPO_DRF_01,      CAMPO_DRF_02,      CAMPO_DRF_03,      CAMPO_DRF_04,"
   g_str_Parame = g_str_Parame & "           CAMPO_DRF_05,      CAMPO_DRF_06,       CAMPO_CAB_01,       CAMPO_CAB_02,       CAMPO_CAB_03,       CAMPO_CAB_04,      CAMPO_CAB_05,      CAMPO_CAB_06,      CAMPO_CAB_07,      CAMPO_CAB_08,"
   g_str_Parame = g_str_Parame & "           CAMPO_CAB_09,      CAMPO_CAB_10,       CAMPO_CAB_11,       CAMPO_CAB_12,       CAMPO_CAB_13,       CAMPO_CAB_14,      CAMPO_CAB_15,      CAMPO_CAB_16,      CAMPO_CAB_17,      CAMPO_CAB_18_1,"
   g_str_Parame = g_str_Parame & "           CAMPO_CAB_18_2,    CAMPO_CAB_19,       CAMPO_CAB_20,       CAMPO_CAB_21,       CAMPO_CAB_22,       CAMPO_CAB_23,      CAMPO_CAB_24,      CAMPO_CAB_25,      CAMPO_CAB_26,      CAMPO_CAB_27,"
   g_str_Parame = g_str_Parame & "           CAMPO_CAB_28_1,    CAMPO_CAB_28_2,     CAMPO_CAB_28_3,     CAMPO_CAB_28_4,     CAMPO_CAB_28_5,     CAMPO_DET_01,      CAMPO_DET_02,      CAMPO_DET_03,      CAMPO_DET_04,"
   g_str_Parame = g_str_Parame & "           MAX(CAMPO_DET_05) AS CAMPO_DET_05,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_06,      CAMPO_DET_07,       CAMPO_DET_08,       CAMPO_DET_09,       CAMPO_DET_10_1,     CAMPO_DET_10_2,    CAMPO_DET_10_3,    CAMPO_DET_10_4,    CAMPO_DET_10_5,    CAMPO_DET_11,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_12,      CAMPO_DET_13,       CAMPO_DET_14,       CAMPO_DET_15_1_1,   CAMPO_DET_15_2_1,   CAMPO_DET_15_1_2,  CAMPO_DET_15_2_2,  CAMPO_DET_15_1_3,  CAMPO_DET_15_2_3,  CAMPO_DET_15_1_4,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_15_2_4,  CAMPO_DET_15_1_5,   CAMPO_DET_15_2_5,   CAMPO_DET_15_1_6,   CAMPO_DET_15_2_6,   CAMPO_DET_15_1_7,  CAMPO_DET_15_2_7,  CAMPO_DET_15_1_8,  CAMPO_DET_15_2_8,  CAMPO_DET_15_1_9,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_15_2_9,  CAMPO_DET_15_1_10,  CAMPO_DET_15_2_10,  CAMPO_DET_15_1_11,  CAMPO_DET_15_2_11,  CAMPO_DET_37,      CAMPO_DET_38,      CAMPO_DET_39_1,    CAMPO_DET_39_2,    CAMPO_DET_39_3,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_39_4,    CAMPO_DET_39_5,     CAMPO_DET2_01,      CAMPO_DET2_02,      CAMPO_DET2_03,      CAMPO_DET2_04,"
   g_str_Parame = g_str_Parame & "           MAX(CAMPO_DET2_05) AS CAMPO_DET2_05,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_06,     CAMPO_DET2_07,      CAMPO_DET2_08,      CAMPO_DET2_09,      CAMPO_DET2_10_1,    CAMPO_DET2_10_2,   CAMPO_DET2_10_3,   CAMPO_DET2_10_4,   CAMPO_DET2_10_5,   CAMPO_DET2_11,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_12,     CAMPO_DET2_13,      CAMPO_DET2_14,      CAMPO_DET2_15_1_1,  CAMPO_DET2_15_2_1,  CAMPO_DET2_15_1_2, CAMPO_DET2_15_2_2, CAMPO_DET2_15_1_3, CAMPO_DET2_15_2_3, CAMPO_DET2_15_1_4,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_15_2_4, CAMPO_DET2_15_1_5,  CAMPO_DET2_15_2_5,  CAMPO_DET2_15_1_6,  CAMPO_DET2_15_2_6,  CAMPO_DET2_15_1_7, CAMPO_DET2_15_2_7, CAMPO_DET2_15_1_8, CAMPO_DET2_15_2_8, CAMPO_DET2_15_1_9,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_15_2_9, CAMPO_DET2_15_1_10, CAMPO_DET2_15_2_10, CAMPO_DET2_15_1_11, CAMPO_DET2_15_2_11, CAMPO_DET2_37,     CAMPO_DET2_38,     CAMPO_DET2_39_1,   CAMPO_DET2_39_2,   CAMPO_DET2_39_3,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_39_4,   CAMPO_DET2_39_5,    CAMPO_ADI_01,       CAMPO_ADI_02,       CAMPO_ADI_03,       CAMPO_ADI_04,      OPERACION,         NUMERO_MOVIMIENTO, FECHA_CANCELACION, SITUACION,"
   g_str_Parame = g_str_Parame & "           FECHA_DEPOSITO"
   g_str_Parame = g_str_Parame & "     FROM ( "
   g_str_Parame = g_str_Parame & "              SELECT 'IDE'                                                                                               AS CAMPO_IDE_01, "
   g_str_Parame = g_str_Parame & "                     'F'                                                                                                 AS CAMPO_IDE_02, "
   g_str_Parame = g_str_Parame & "                     SUBSTR(CAJMOV_FECDEP,1,4) || '-' || SUBSTR(CAJMOV_FECDEP,5,2) || '-' || SUBSTR(CAJMOV_FECDEP,7,2)   AS CAMPO_IDE_03, "
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_IDE_04,"
   g_str_Parame = g_str_Parame & "                     '01'                                                                                                AS CAMPO_IDE_05,              " '--CATALOGO N°1.- FACTURA
   g_str_Parame = g_str_Parame & "                     C.CATSUN_CODIGO                                                                                     AS CAMPO_IDE_06,              " '--CATALOGO N°2.-
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_IDE_07,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_IDE_08,"
   
   g_str_Parame = g_str_Parame & "                     'EMI'                                                                                               AS CAMPO_EMI_01,"
   g_str_Parame = g_str_Parame & "                     'F'                                                                                                 AS CAMPO_EMI_02,"
   g_str_Parame = g_str_Parame & "                     '6'                                                                                                 AS CAMPO_EMI_03,              " '--CATALOGO N°6.-
   g_str_Parame = g_str_Parame & "                     '20511904162'                                                                                       AS CAMPO_EMI_04,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_05,"
   g_str_Parame = g_str_Parame & "                     'EDPYME MICASITA SA'                                                                                AS CAMPO_EMI_06,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_07,              " '--ATALOGO N°13.-UBIGEO
   g_str_Parame = g_str_Parame & "                     'AV RIVERA NAVARRETE 645'                                                                           AS CAMPO_EMI_08,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_09,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_10,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_11,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_12,"
   g_str_Parame = g_str_Parame & "                     'PE'                                                                                                AS CAMPO_EMI_13,              " '--CATALOGO N°4.-
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_14,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_15,"
   g_str_Parame = g_str_Parame & "                     '0000'                                                                                              AS CAMPO_EMI_16,"
   
   g_str_Parame = g_str_Parame & "                     'REC'                                                                                               AS CAMPO_REC_01,"
   g_str_Parame = g_str_Parame & "                     'F'                                                                                                 AS CAMPO_REC_02,"
   g_str_Parame = g_str_Parame & "                     TRIM(A.CAJMOV_TIPDOC)                                                                               AS CAMPO_REC_03,              " '--CATALOGO N°6.-
   g_str_Parame = g_str_Parame & "                     TRIM(A.CAJMOV_NUMDOC)                                                                               AS CAMPO_REC_04,"
   g_str_Parame = g_str_Parame & "                     TRIM(D.DATGEN_APEPAT) || ' ' || TRIM(D.DATGEN_APEMAT) || ' ' || TRIM(D.DATGEN_NOMBRE)               AS CAMPO_REC_05,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_REC_06,"
   g_str_Parame = g_str_Parame & "                     TRIM(H.PARDES_DESCRI)                                                                               AS CAMPO_REC_07,"
   g_str_Parame = g_str_Parame & "                     TRIM(I.PARDES_DESCRI)                                                                               AS CAMPO_REC_08,"
   g_str_Parame = g_str_Parame & "                     TRIM(G.PARDES_DESCRI)                                                                               AS CAMPO_REC_09,"
   g_str_Parame = g_str_Parame & "                     'PE'                                                                                                AS CAMPO_REC_10,"
   g_str_Parame = g_str_Parame & "                     TRIM(D.DATGEN_TELEFO)                                                                               AS CAMPO_REC_11,"
   g_str_Parame = g_str_Parame & "                     TRIM(D.DATGEN_DIRELE)                                                                               AS CAMPO_REC_12,"
   
   g_str_Parame = g_str_Parame & "                     'DRF'                                                                                               AS CAMPO_DRF_01,"
   g_str_Parame = g_str_Parame & "                     'F'                                                                                                 AS CAMPO_DRF_02,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DRF_03,               "
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DRF_04,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DRF_05,               " '--PARA NOTA CREDITO/DEBITO
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DRF_06,"
   
   g_str_Parame = g_str_Parame & "                     'CAB'                                                                                               AS CAMPO_CAB_01,"
   g_str_Parame = g_str_Parame & "                     'F'                                                                                                 AS CAMPO_CAB_02,"
   g_str_Parame = g_str_Parame & "                     '1001'                                                                                              AS CAMPO_CAB_03,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_CAB_04,"
   g_str_Parame = g_str_Parame & "                     '1002'                                                                                              AS CAMPO_CAB_05,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "                     A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_06,"
   g_str_Parame = g_str_Parame & "                     '1003'                                                                                              AS CAMPO_CAB_07,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "                     '0.00'                                                                                              AS CAMPO_CAB_08,"
   g_str_Parame = g_str_Parame & "                     '1004'                                                                                              AS CAMPO_CAB_09,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "                     '0.00'                                                                                              AS CAMPO_CAB_10,"
   g_str_Parame = g_str_Parame & "                     '1000'                                                                                              AS CAMPO_CAB_11,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_CAB_12,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_13,"
   
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_14,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_15,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_16,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_17,"
   g_str_Parame = g_str_Parame & "                     '1000'                                                                                              AS CAMPO_CAB_18_1,"
   g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_CAB_18_2,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_19,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_20,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_21,"
   g_str_Parame = g_str_Parame & "                     A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_22,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_23,                " '23.1 , 23.2 , 23.3 , 23.4 , 23.5 , 23.6 , 23.7 , 23.8
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_24,"
   g_str_Parame = g_str_Parame & "                     '13'                                                                                                AS CAMPO_CAB_25,                " '                 --CATALOGO N°17.-
   g_str_Parame = g_str_Parame & "                     '[1000'                                                                                             AS CAMPO_CAB_26,                " '26.1 , 26.2      --CATALOGO N°15.- DETALLE EN LETRAS DEL IMPORTE
   g_str_Parame = g_str_Parame & "                     0.0                                                                                                 AS CAMPO_CAB_27,"
   g_str_Parame = g_str_Parame & "                     'false'                                                                                             AS CAMPO_CAB_28_1,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_28_2,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_28_3,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_28_4,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_28_5,"
   
   g_str_Parame = g_str_Parame & "                     'DET1'                                                                                              AS CAMPO_DET_01,"
   g_str_Parame = g_str_Parame & "                     'F'                                                                                                 AS CAMPO_DET_02,"
   g_str_Parame = g_str_Parame & "                     '001'                                                                                               AS CAMPO_DET_03,                " '-- Número de orden de ítem
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_04,                "
   g_str_Parame = g_str_Parame & "                     'INTERES ' ||"
   g_str_Parame = g_str_Parame & "                     SUBSTR(TO_CHAR(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(HIPCUO_FECVCT,1,4)  AS CAMPO_DET_05," 'CAJMOV_FECDEP
   g_str_Parame = g_str_Parame & "                     1.000                                                                                               AS CAMPO_DET_06,"
   g_str_Parame = g_str_Parame & "                     'NIU'                                                                                               AS CAMPO_DET_07,                " '--CATALOGO N°3.-
   g_str_Parame = g_str_Parame & "                     A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_08,"
   g_str_Parame = g_str_Parame & "                     1.000 * A.CAJMOV_INTERE                                                                             AS CAMPO_DET_09,"
   g_str_Parame = g_str_Parame & "                     '1000'                                                                                              AS CAMPO_DET_10_1,              " '--CATALOGO N°5,7 u 8.-
   g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_DET_10_2,"
   g_str_Parame = g_str_Parame & "                     '30'                                                                                                AS CAMPO_DET_10_3,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_10_4,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_10_5,"
   g_str_Parame = g_str_Parame & "                     A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_11,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_12,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_13,"
   g_str_Parame = g_str_Parame & "                     '84121901'                                                                                          AS CAMPO_DET_14,                " '--CATALOGO N°15.-
   g_str_Parame = g_str_Parame & "                     '7001'                                                                                              AS CAMPO_DET_15_1_1,"
   g_str_Parame = g_str_Parame & "                     '1'                                                                                                 AS CAMPO_DET_15_2_1,                " '--CATALOGO N°26.- Si es construcción /adquisición
   g_str_Parame = g_str_Parame & "                     '7002'                                                                                              AS CAMPO_DET_15_1_2,"
   g_str_Parame = g_str_Parame & "                     CASE WHEN L.SOLMAE_PRIVIV = 1 THEN 3  "
   g_str_Parame = g_str_Parame & "                          WHEN L.SOLMAE_PRIVIV = 2 THEN 0 END                                                            AS CAMPO_DET_15_2_2,                " '--CATALOGO N°27.- VERIFICAR
   g_str_Parame = g_str_Parame & "                     '7003'                                                                                              AS CAMPO_DET_15_1_3,"
   g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_3,"
   g_str_Parame = g_str_Parame & "                     '7004'                                                                                              AS CAMPO_DET_15_1_4,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_15_2_4,                " '--NRO CONTRATO O PRESTAMO
   g_str_Parame = g_str_Parame & "                     '7005'                                                                                              AS CAMPO_DET_15_1_5,"
   g_str_Parame = g_str_Parame & "                     SUBSTR(J.HIPMAE_FECACT,1,4) || '-' || SUBSTR(J.HIPMAE_FECACT,5,2) || '-' || SUBSTR(J.HIPMAE_FECACT,7,2)"
   g_str_Parame = g_str_Parame & "                                                                                                                         AS CAMPO_DET_15_2_5,"
   g_str_Parame = g_str_Parame & "                     '7006'                                                                                              AS CAMPO_DET_15_1_6,"
   g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_6,"                 '/*P.SOLINM_UBIGEO*/
   g_str_Parame = g_str_Parame & "                     '7007'                                                                                              AS CAMPO_DET_15_1_7,"
   g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_7,"
   g_str_Parame = g_str_Parame & "                     '7008'                                                                                              AS CAMPO_DET_15_1_8,"
   g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_8,"
   g_str_Parame = g_str_Parame & "                     '7009'                                                                                              AS CAMPO_DET_15_1_9,"
   g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_9,"
   g_str_Parame = g_str_Parame & "                     '7010'                                                                                              AS CAMPO_DET_15_1_10,"
   g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_10,"
   g_str_Parame = g_str_Parame & "                     '7011'                                                                                              AS CAMPO_DET_15_1_11,"
   g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_11,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_37,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_38,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_39_1,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_39_2,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_39_3,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_39_4,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_39_5,"
   
   g_str_Parame = g_str_Parame & "                     'DET2'                                                                                              AS CAMPO_DET2_01,"
   g_str_Parame = g_str_Parame & "                     'F'                                                                                                 AS CAMPO_DET2_02,"
   g_str_Parame = g_str_Parame & "                     '002'                                                                                               AS CAMPO_DET2_03,                " '-- Número de orden de ítem
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_04,                "
   g_str_Parame = g_str_Parame & "                     'OTROS IMPORTES ' ||"
   g_str_Parame = g_str_Parame & "                     SUBSTR(TO_CHAR(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(HIPCUO_FECVCT,1,4)  AS CAMPO_DET2_05," 'CAJMOV_FECDEP
   g_str_Parame = g_str_Parame & "                     1.000                                                                                               AS CAMPO_DET2_06,"
   g_str_Parame = g_str_Parame & "                     'NIU'                                                                                               AS CAMPO_DET2_07,                " '--CATALOGO N°3.-
   g_str_Parame = g_str_Parame & "                     (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)                                                                 AS CAMPO_DET2_08,"
   g_str_Parame = g_str_Parame & "                     1.000 * (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)                                                         AS CAMPO_DET2_09,"
   g_str_Parame = g_str_Parame & "                     '1000'                                                                                              AS CAMPO_DET2_10_1,                 " '--CATALOGO N°5,7 u 8.-
   g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_DET2_10_2,"
   g_str_Parame = g_str_Parame & "                     '30'                                                                                                AS CAMPO_DET2_10_3,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_10_4,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_10_5,"
   g_str_Parame = g_str_Parame & "                     A.CAJMOV_IMPPAG - A.CAJMOV_INTERE                                                                   AS CAMPO_DET2_11,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_12,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_13,"
   g_str_Parame = g_str_Parame & "                     '84121501'                                                                                          AS CAMPO_DET2_14,                    " '--CATALOGO N°15.-
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_1,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_1,                " '--NRO CONTRATO O PRESTAMO
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_2,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_2,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_3,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_3,                " '--CATALOGO N°26.- Si es construcción /adquisición
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_4,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_4,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_5,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_5,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_6,                " '--CATALOGO N°27.- VERIFICAR
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_6,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_7,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_7,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_8,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_8,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_9,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_9,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_10,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_10,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_11,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_11,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_37,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_38,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_39_1,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_39_2,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_39_3,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_39_4,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_39_5,"

   g_str_Parame = g_str_Parame & "                     'ADI1'                                                                                              AS CAMPO_ADI_01,"
   g_str_Parame = g_str_Parame & "                     'F'                                                                                                 AS CAMPO_ADI_02,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_ADI_03,"
   g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_ADI_04,"
   g_str_Parame = g_str_Parame & "                     TRIM(A.CAJMOV_NUMOPE)                                                                               AS OPERACION, "
   g_str_Parame = g_str_Parame & "                     CAJMOV_NUMMOV                                                                                       AS NUMERO_MOVIMIENTO , "
   g_str_Parame = g_str_Parame & "                     HIPMAE_FECCAN                                                                                       AS FECHA_CANCELACION ,"
   g_str_Parame = g_str_Parame & "                     HIPMAE_SITUAC                                                                                       AS SITUACION, "
   g_str_Parame = g_str_Parame & "                     CAJMOV_FECDEP                                                                                       AS FECHA_DEPOSITO "
   g_str_Parame = g_str_Parame & "                FROM OPE_CAJMOV A "
   g_str_Parame = g_str_Parame & "                     INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '204' AND B.PARDES_CODITE = A.CAJMOV_MONPAG "
   g_str_Parame = g_str_Parame & "                     INNER JOIN CNTBL_CATSUN C ON C.CATSUN_NROCAT = 2 AND C.CATSUN_DESCRI = TRIM(B.PARDES_DESCRI) "
   g_str_Parame = g_str_Parame & "                     INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = A.CAJMOV_TIPDOC AND D.DATGEN_NUMDOC = A.CAJMOV_NUMDOC "
   g_str_Parame = g_str_Parame & "                     INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 201 AND E.PARDES_CODITE = D.DATGEN_TIPVIA "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 202 AND F.PARDES_CODITE = D.DATGEN_TIPZON "
   g_str_Parame = g_str_Parame & "                     INNER JOIN MNT_PARDES G ON G.PARDES_CODGRP = 101 AND G.PARDES_CODITE = D.DATGEN_UBIGEO "
   g_str_Parame = g_str_Parame & "                     INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 101 AND H.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,2)||'0000' "
   g_str_Parame = g_str_Parame & "                     INNER JOIN MNT_PARDES I ON I.PARDES_CODGRP = 101 AND I.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,4)||'00' "
   g_str_Parame = g_str_Parame & "                     INNER JOIN CRE_HIPMAE J ON J.HIPMAE_NUMOPE = A.CAJMOV_NUMOPE "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN CRE_HIPGAR K ON K.HIPGAR_NUMOPE = J.HIPMAE_NUMOPE AND K.HIPGAR_BIEGAR = 1 "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN CRE_SOLMAE L ON L.SOLMAE_NUMERO = J.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN CRE_SOLINM P ON P.SOLINM_NUMSOL = J.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES Q ON Q.PARDES_CODGRP = 201 AND Q.PARDES_CODITE = P.SOLINM_TIPVIA "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES R ON R.PARDES_CODGRP = 202 AND R.PARDES_CODITE = P.SOLINM_TIPZON "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES S ON S.PARDES_CODGRP = 101 AND S.PARDES_CODITE = P.SOLINM_UBIGEO "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES T ON T.PARDES_CODGRP = 101 AND T.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,4)||'00' "
   g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES U ON U.PARDES_CODGRP = 101 AND U.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,2)||'0000' "
   g_str_Parame = g_str_Parame & "                     INNER JOIN CRE_HIPPAG V ON V.HIPPAG_NUMOPE = A.CAJMOV_NUMOPE AND V.HIPPAG_FECPAG = A.CAJMOV_FECDEP AND V.HIPPAG_NUMMOV = A.CAJMOV_NUMMOV"
   g_str_Parame = g_str_Parame & "                     INNER JOIN CRE_HIPCUO W ON W.HIPCUO_NUMOPE = V.HIPPAG_NUMOPE AND W.HIPCUO_TIPCRO = 1 AND W.HIPCUO_NUMCUO = V.HIPPAG_NUMCUO"
   g_str_Parame = g_str_Parame & "               WHERE CAJMOV_SUCMOV IS NOT NULL "
   g_str_Parame = g_str_Parame & "                 AND CAJMOV_USUMOV IS NOT NULL "
   g_str_Parame = g_str_Parame & "                 AND CAJMOV_FECMOV > 0"
   g_str_Parame = g_str_Parame & "                 AND CAJMOV_NUMMOV > 0 "
   g_str_Parame = g_str_Parame & "                 AND CAJMOV_CODBAN IS NOT NULL "
   
   If Chk_FecAct.Value = 0 Then
      g_str_Parame = g_str_Parame & "              AND CAJMOV_FECDEP = '" & l_str_FecCar & "' "
   Else
      g_str_Parame = g_str_Parame & "              AND CAJMOV_FECDEP <= '" & l_str_FecCar & "' "
   End If
   
   g_str_Parame = g_str_Parame & "                 AND CAJMOV_FLGPRO = 0 "
   g_str_Parame = g_str_Parame & "                 AND CAJMOV_TIPMOV = '1102' "
   g_str_Parame = g_str_Parame & "                 AND CAJMOV_TIPDOC IN (1, 6) "                      '1-DNI y 6-RUC
   g_str_Parame = g_str_Parame & "               ORDER BY A.CAJMOV_FECMOV , A.CAJMOV_NUMMOV "
   g_str_Parame = g_str_Parame & "      )"
   g_str_Parame = g_str_Parame & "      GROUP BY  CAMPO_IDE_01,      CAMPO_IDE_02,       CAMPO_IDE_03,       CAMPO_IDE_04,       CAMPO_IDE_05,       CAMPO_IDE_06,      CAMPO_IDE_07,      CAMPO_IDE_08,      CAMPO_EMI_01,      CAMPO_EMI_02,"
   g_str_Parame = g_str_Parame & "                CAMPO_EMI_03,      CAMPO_EMI_04,       CAMPO_EMI_05,       CAMPO_EMI_06,       CAMPO_EMI_07,       CAMPO_EMI_08,      CAMPO_EMI_09,      CAMPO_EMI_10,      CAMPO_EMI_11,      CAMPO_EMI_12,"
   g_str_Parame = g_str_Parame & "                CAMPO_EMI_13,      CAMPO_EMI_14,       CAMPO_EMI_15,       CAMPO_EMI_16,       CAMPO_REC_01,       CAMPO_REC_02,      CAMPO_REC_03,      CAMPO_REC_04,      CAMPO_REC_05,      CAMPO_REC_06,"
   g_str_Parame = g_str_Parame & "                CAMPO_REC_07,      CAMPO_REC_08,       CAMPO_REC_09,       CAMPO_REC_10,       CAMPO_REC_11,       CAMPO_REC_12,      CAMPO_DRF_01,      CAMPO_DRF_02,      CAMPO_DRF_03,      CAMPO_DRF_04,"
   g_str_Parame = g_str_Parame & "                CAMPO_DRF_05,      CAMPO_DRF_06,       CAMPO_CAB_01,       CAMPO_CAB_02,       CAMPO_CAB_03,       CAMPO_CAB_04,      CAMPO_CAB_05,      CAMPO_CAB_06,      CAMPO_CAB_07,      CAMPO_CAB_08,"
   g_str_Parame = g_str_Parame & "                CAMPO_CAB_09,      CAMPO_CAB_10,       CAMPO_CAB_11,       CAMPO_CAB_12,       CAMPO_CAB_13,       CAMPO_CAB_14,      CAMPO_CAB_15,      CAMPO_CAB_16,      CAMPO_CAB_17,      CAMPO_CAB_18_1,"
   g_str_Parame = g_str_Parame & "                CAMPO_CAB_18_2,    CAMPO_CAB_19,       CAMPO_CAB_20,       CAMPO_CAB_21,       CAMPO_CAB_22,       CAMPO_CAB_23,      CAMPO_CAB_24,      CAMPO_CAB_25,      CAMPO_CAB_26,      CAMPO_CAB_27,"
   g_str_Parame = g_str_Parame & "                CAMPO_CAB_28_1,    CAMPO_CAB_28_2,     CAMPO_CAB_28_3,     CAMPO_CAB_28_4,     CAMPO_CAB_28_5,     CAMPO_DET_01,      CAMPO_DET_02,      CAMPO_DET_03,      CAMPO_DET_04,"
   '                                                CAMPO_DET_05,
   g_str_Parame = g_str_Parame & "                CAMPO_DET_06,      CAMPO_DET_07,       CAMPO_DET_08,       CAMPO_DET_09,       CAMPO_DET_10_1,     CAMPO_DET_10_2,    CAMPO_DET_10_3,    CAMPO_DET_10_4,    CAMPO_DET_10_5,    CAMPO_DET_11,"
   g_str_Parame = g_str_Parame & "                CAMPO_DET_12,      CAMPO_DET_13,       CAMPO_DET_14,       CAMPO_DET_15_1_1,   CAMPO_DET_15_2_1,   CAMPO_DET_15_1_2,  CAMPO_DET_15_2_2,  CAMPO_DET_15_1_3,  CAMPO_DET_15_2_3,  CAMPO_DET_15_1_4,"
   g_str_Parame = g_str_Parame & "                CAMPO_DET_15_2_4,  CAMPO_DET_15_1_5,   CAMPO_DET_15_2_5,   CAMPO_DET_15_1_6,   CAMPO_DET_15_2_6,   CAMPO_DET_15_1_7,  CAMPO_DET_15_2_7,  CAMPO_DET_15_1_8,  CAMPO_DET_15_2_8,  CAMPO_DET_15_1_9,"
   g_str_Parame = g_str_Parame & "                CAMPO_DET_15_2_9,  CAMPO_DET_15_1_10,  CAMPO_DET_15_2_10,  CAMPO_DET_15_1_11,  CAMPO_DET_15_2_11,  CAMPO_DET_37,      CAMPO_DET_38,      CAMPO_DET_39_1,    CAMPO_DET_39_2,    CAMPO_DET_39_3,"
   g_str_Parame = g_str_Parame & "                CAMPO_DET_39_4,    CAMPO_DET_39_5,     CAMPO_DET2_01,      CAMPO_DET2_02,      CAMPO_DET2_03,      CAMPO_DET2_04,"
   '                                                CAMPO_DET2_05,
   g_str_Parame = g_str_Parame & "                CAMPO_DET2_06,     CAMPO_DET2_07,      CAMPO_DET2_08,      CAMPO_DET2_09,      CAMPO_DET2_10_1,    CAMPO_DET2_10_2,   CAMPO_DET2_10_3,   CAMPO_DET2_10_4,   CAMPO_DET2_10_5,   CAMPO_DET2_11,"
   g_str_Parame = g_str_Parame & "                CAMPO_DET2_12,     CAMPO_DET2_13,      CAMPO_DET2_14,      CAMPO_DET2_15_1_1,  CAMPO_DET2_15_2_1,  CAMPO_DET2_15_1_2, CAMPO_DET2_15_2_2, CAMPO_DET2_15_1_3, CAMPO_DET2_15_2_3, CAMPO_DET2_15_1_4,"
   g_str_Parame = g_str_Parame & "                CAMPO_DET2_15_2_4, CAMPO_DET2_15_1_5,  CAMPO_DET2_15_2_5,  CAMPO_DET2_15_1_6,  CAMPO_DET2_15_2_6,  CAMPO_DET2_15_1_7, CAMPO_DET2_15_2_7, CAMPO_DET2_15_1_8, CAMPO_DET2_15_2_8, CAMPO_DET2_15_1_9,"
   g_str_Parame = g_str_Parame & "                CAMPO_DET2_15_2_9, CAMPO_DET2_15_1_10, CAMPO_DET2_15_2_10, CAMPO_DET2_15_1_11, CAMPO_DET2_15_2_11, CAMPO_DET2_37,     CAMPO_DET2_38,     CAMPO_DET2_39_1,   CAMPO_DET2_39_2,   CAMPO_DET2_39_3,"
   g_str_Parame = g_str_Parame & "                CAMPO_DET2_39_4,   CAMPO_DET2_39_5,    CAMPO_ADI_01,       CAMPO_ADI_02,       CAMPO_ADI_03,       CAMPO_ADI_04,      OPERACION,         NUMERO_MOVIMIENTO, FECHA_CANCELACION, SITUACION,"
   g_str_Parame = g_str_Parame & "                FECHA_DEPOSITO"
   

'   g_str_Parame = g_str_Parame & "      GROUP BY CAMPO_IDE_01 , CAMPO_IDE_02  , CAMPO_IDE_03 , CAMPO_IDE_04 , CAMPO_IDE_05   , CAMPO_IDE_06   , CAMPO_IDE_07   , CAMPO_IDE_08  , CAMPO_EMI_01  , CAMPO_EMI_02,"
'   g_str_Parame = g_str_Parame & "               CAMPO_EMI_03 , CAMPO_EMI_04  , CAMPO_EMI_05 , CAMPO_EMI_06 , CAMPO_EMI_07   , CAMPO_EMI_08   , CAMPO_EMI_09   , CAMPO_EMI_10  , CAMPO_EMI_11  , CAMPO_EMI_12,"
'   g_str_Parame = g_str_Parame & "               CAMPO_EMI_13 , CAMPO_EMI_14  , CAMPO_EMI_15 , CAMPO_REC_01 , CAMPO_REC_02   , CAMPO_REC_03   , CAMPO_REC_04   , CAMPO_REC_05  , CAMPO_REC_06  , CAMPO_REC_07,"
'   g_str_Parame = g_str_Parame & "               CAMPO_REC_08 , CAMPO_REC_09  , CAMPO_REC_10 , CAMPO_REC_11 , CAMPO_REC_12   , CAMPO_DRF_01   , CAMPO_DRF_02   , CAMPO_DRF_03  , CAMPO_DRF_04  , CAMPO_DRF_05,"
'   g_str_Parame = g_str_Parame & "               CAMPO_DRF_06 , CAMPO_CAB_01  , CAMPO_CAB_02 , CAMPO_CAB_03 , CAMPO_CAB_04   , CAMPO_CAB_05   , CAMPO_CAB_06   , CAMPO_CAB_07  , CAMPO_CAB_08  , CAMPO_CAB_09,"
'   g_str_Parame = g_str_Parame & "               CAMPO_CAB_10 , CAMPO_CAB_11  , CAMPO_CAB_12 , CAMPO_CAB_13 , CAMPO_CAB_14   , CAMPO_CAB_15   , CAMPO_CAB_16   , CAMPO_CAB_17  , CAMPO_CAB_18_1, CAMPO_CAB_18_2,"
'   g_str_Parame = g_str_Parame & "               CAMPO_CAB_19 , CAMPO_CAB_20  , CAMPO_CAB_21 , CAMPO_CAB_22 , CAMPO_CAB_23   , CAMPO_CAB_24   , CAMPO_CAB_25   , CAMPO_CAB_26  , CAMPO_CAB_27  , CAMPO_DET_01,"
'   g_str_Parame = g_str_Parame & "               CAMPO_DET_02 , CAMPO_DET_03  , CAMPO_DET_04 , "
'   g_str_Parame = g_str_Parame & "               CAMPO_DET_06 , CAMPO_DET_07  , CAMPO_DET_08 , CAMPO_DET_09 , CAMPO_DET_10_1 , CAMPO_DET_10_2 , CAMPO_DET_10_3 , CAMPO_DET_11  , CAMPO_DET_12  , CAMPO_DET_13,"
'   g_str_Parame = g_str_Parame & "               CAMPO_DET_14 , CAMPO_DET_15  , CAMPO_DET_16 , CAMPO_DET_17 , CAMPO_DET_18   , CAMPO_DET_19   , CAMPO_DET_20   , CAMPO_DET_21  , CAMPO_DET_22  , CAMPO_DET_23,"
'   g_str_Parame = g_str_Parame & "               CAMPO_DET_24 , CAMPO_DET_25  , CAMPO_DET_26 , CAMPO_DET_27 , CAMPO_DET_28   , CAMPO_DET_29   , CAMPO_DET_30   , CAMPO_DET_31  , CAMPO_DET_32  , CAMPO_DET_33,"
'   g_str_Parame = g_str_Parame & "               CAMPO_DET_34 , CAMPO_DET_35  , CAMPO_DET_36 , CAMPO_DET_37 , CAMPO_DET2_01  , CAMPO_DET2_02  , CAMPO_DET2_03  , CAMPO_DET2_04 , "
'   g_str_Parame = g_str_Parame & "               CAMPO_DET2_06, CAMPO_DET2_07, CAMPO_DET2_08 , CAMPO_DET2_09, CAMPO_DET2_10_1, CAMPO_DET2_10_2, CAMPO_DET2_10_3, CAMPO_DET2_11 , CAMPO_DET2_12 , CAMPO_DET2_13,"
'   g_str_Parame = g_str_Parame & "               CAMPO_DET2_14, CAMPO_DET2_15, CAMPO_DET2_16 , CAMPO_DET2_17, CAMPO_DET2_18  , CAMPO_DET2_19  , CAMPO_DET2_20  , CAMPO_DET2_21 , CAMPO_DET2_22 , CAMPO_DET2_23,"
'   g_str_Parame = g_str_Parame & "               CAMPO_DET2_24, CAMPO_DET2_25, CAMPO_DET2_26 , CAMPO_DET2_27, CAMPO_DET2_28  , CAMPO_DET2_29  , CAMPO_DET2_30  , CAMPO_DET2_31 , CAMPO_DET2_32 , CAMPO_DET2_33,"
'   g_str_Parame = g_str_Parame & "               CAMPO_DET2_34, CAMPO_DET2_35, CAMPO_DET2_36 , CAMPO_DET2_37, CAMPO_ADI_01   , CAMPO_ADI_02   , CAMPO_ADI_03   , CAMPO_ADI_04  , OPERACION     , NUMERO_MOVIMIENTO,"
'   g_str_Parame = g_str_Parame & "               FECHA_CANCELACION , SITUACION, FECHA_DEPOSITO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta principal, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontró ningún registro.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro anterior en OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
    
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
      moddat_g_str_NumOpe = g_rst_Princi!OPERACION
      moddat_g_str_Codigo = g_rst_Princi!NUMERO_MOVIMIENTO
      
      If g_rst_Princi!SITUACION <> 2 Then
         If g_rst_Princi!FECHA_DEPOSITO <= g_rst_Princi!FECHA_CANCELACION Then
            GoTo Ingresar
         End If
      Else
      
Ingresar:
         Call fs_Obtener_Codigo("01", r_lng_Contad, r_int_SerFac, r_lng_NumFac)
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELE (      "
         g_str_Parame = g_str_Parame & " DOCELE_CODIGO                 , "
         g_str_Parame = g_str_Parame & " DOCELE_NUMOPE                 , "
         g_str_Parame = g_str_Parame & " DOCELE_NUMMOV                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FECPRO                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FECAUT                 , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_FECEMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_HOREMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPMON             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_NUMORC             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_FECVCT             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_NOMCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_UBIGEO             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_URBANI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_TELEMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_COREMI             , "
         
         g_str_Parame = g_str_Parame & " DOCELE_EMI_CODSUN             , "
         
         g_str_Parame = g_str_Parame & " DOCELE_REC_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_TELREC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_CORREC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_CODMOT             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_DESMOT             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXP      , "
         
'         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_INDICA_CARDSC      , "
         
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_REGPER      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_BASIMP_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOPER             , "
         
'         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOTOT_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_FACTOR_TASPER      , "
         
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIMP             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOIMP             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_OTRCAR             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_TOTDSC      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTDSC             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_IMPTOT_DOCUME      , "

'         g_str_Parame = g_str_Parame & " DOCELE_CAB_DSCGLO             , "

         g_str_Parame = g_str_Parame & " DOCELE_CAB_INFPPG             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTANT             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TIPOPE             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_LEYEND             , "
         
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOTOT_IMP         , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CARDSC             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODMOT_CARDS       , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_FACTOR_CARDS       , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MONTO_CARDSC       , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOBAS_CARDSC      , "
         
         g_str_Parame = g_str_Parame & " DOCELE_ADI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_TITADI             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_VALADI             , "
         g_str_Parame = g_str_Parame & " DOCELE_FLGENV                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FLGRPT                 , "
         g_str_Parame = g_str_Parame & " DOCELE_SITUAC                 , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES ( "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & " , "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
         g_str_Parame = g_str_Parame & "" & moddat_g_str_Codigo & " , "
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "     , "
         g_str_Parame = g_str_Parame & " NULL, "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_05 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_06 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_05 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_06 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_09 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_10 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_11 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_12 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_13 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_14 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_15 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_16 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_05 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_REC_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "'" & Mid(Replace(g_rst_Princi!CAMPO_REC_06, "  ", " "), 1, 100) & "'                                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_09 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_10 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_11 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_12 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_04 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DRF_05) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DRF_05 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_06 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_03 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_04) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_04 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_06 & "                    , "
         End If
               
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_08 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_09 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_10) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_10 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_11 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_12 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_13 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_14 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_15) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_15 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_16) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_16 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_17) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_17 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_18_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_18_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_18_2 & "                  , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_19) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_19 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_20 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_21 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_22) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_22 & "                    , "
         End If
         
'         If IsNull(g_rst_Princi!CAMPO_CAB_23) Then
'            g_str_Parame = g_str_Parame & " NULL, "
'         Else
'            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_23 & "                    , "
'         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_23 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_24) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_24 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_25 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_26 & "'                     , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_27 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_28_1 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_28_2 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_28_3 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_28_4 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_28_5 & "'                   , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_04 & "'                     , "
         g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
         g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
         g_str_Parame = g_str_Parame & "" & 1 & "                                               , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                               
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELE, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
   
         
         ''INTERES COMPENSATORIO
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOBAS_IMP      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TASTRI          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         
'         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "

         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD_GS1      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOTOT_IMP      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_INDCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FACCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOBAS_CAR      , "
         
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES (                        "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_04 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_06 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_08 & "                    , "
         End If
   
         If IsNull(g_rst_Princi!CAMPO_DET_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_09 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_10_2 & "                  , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_3 & "'                   , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_4 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_5 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_11 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_12 & "                    , "
         End If
      
'         If IsNull(g_rst_Princi!CAMPO_DET_13) Then
'            g_str_Parame = g_str_Parame & " NULL, "
'         Else
'            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_13 & "                    , "
'         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_13 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_14 & "'                     , "

         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_1 & "'                 , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_15_2_1) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_15_2_1 & "                , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_2 & "'                 , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_15_2_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_15_2_2 & "                , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_3 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_3 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_4 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_4 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_5 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_5 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_6 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_6 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_7 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_7 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_8 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_8 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_9 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_9 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_10 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_10 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_11 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_11 & "'                , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_37 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_38 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_39_1 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_39_2 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_39_3 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_39_4 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_39_5 & "'                   , "
         
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "'                           , "
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de INTERES COMPENSATORIO, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
         
                                      
         ''OTROS IMPORTES
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOBAS_IMP      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TASTRI          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         
'         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "

         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD_GS1      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOTOT_IMP      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_INDCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FACCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOBAS_CAR      , "
         
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES (                        "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_04 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_06 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_08 & "                    , "
         End If
   
         If IsNull(g_rst_Princi!CAMPO_DET2_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_09 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_10_2 & "                  , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_3 & "'                   , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_4 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_5 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_11 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_12 & "                    , "
         End If
      
'         If IsNull(g_rst_Princi!CAMPO_DET2_13) Then
'            g_str_Parame = g_str_Parame & " NULL, "
'         Else
'            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_13 & "                    , "
'         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_13 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_14 & "'                     , "

         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_1 & "'                 , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_15_2_1) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_15_2_1 & "                , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_2 & "'                 , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_15_2_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_15_2_2 & "                , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_3 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_3 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_4 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_4 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_5 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_5 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_6 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_6 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_7 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_7 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_8 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_8 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_9 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_9 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_10 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_10 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_11 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_11 & "'                , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_37 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_38 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_39_1 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_39_2 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_39_3 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_39_4 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_39_5 & "'                   , "
         
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de OTROS IMPORTES, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
                                                                                                   
         DoEvents: DoEvents: DoEvents
            
         'ACTUALIZA EL CAMPO CAJMOV_FLGPRO PARA IDENTIFICAR CUALES SE HAN PROCESADO Y YA SE ENCUENTRAN EN CNTBL_DOCELE
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "UPDATE OPE_CAJMOV SET CAJMOV_FLGPRO = 1 "
         g_str_Parame = g_str_Parame & " WHERE CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' "
         g_str_Parame = g_str_Parame & "   AND CAJMOV_NUMMOV = '" & moddat_g_str_Codigo & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar CAJMOV_FLGPRO de la tabla OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
      
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Genera_FactAnterior")

End Sub
Private Sub fs_Generar_Boletas_NUEVO_FORMATO()
Dim r_lng_Contad     As Long
Dim r_int_SerFac     As Integer
Dim r_lng_NumFac     As Long

   On Error GoTo MyError
   
   Screen.MousePointer = 11
   
   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "   SELECT  CAMPO_IDE_01, CAMPO_IDE_02  , CAMPO_IDE_03 , CAMPO_IDE_04 , CAMPO_IDE_05   , CAMPO_IDE_06   , CAMPO_IDE_07   , CAMPO_IDE_08 , CAMPO_EMI_01  , CAMPO_EMI_02,"
'   g_str_Parame = g_str_Parame & "           CAMPO_EMI_03, CAMPO_EMI_04  , CAMPO_EMI_05 , CAMPO_EMI_06 , CAMPO_EMI_07   , CAMPO_EMI_08   , CAMPO_EMI_09   , CAMPO_EMI_10 , CAMPO_EMI_11  , CAMPO_EMI_12,"
'   g_str_Parame = g_str_Parame & "           CAMPO_EMI_13, CAMPO_EMI_14  , CAMPO_EMI_15 , CAMPO_REC_01 , CAMPO_REC_02   , CAMPO_REC_03   , CAMPO_REC_04   , CAMPO_REC_05 , CAMPO_REC_06  , CAMPO_REC_07,"
'   g_str_Parame = g_str_Parame & "           CAMPO_REC_08, CAMPO_REC_09  , CAMPO_REC_10 , CAMPO_REC_11 , CAMPO_REC_12   , CAMPO_DRF_01   , CAMPO_DRF_02   , CAMPO_DRF_03 , CAMPO_DRF_04  , CAMPO_DRF_05,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DRF_06, CAMPO_CAB_01  , CAMPO_CAB_02 , CAMPO_CAB_03 , CAMPO_CAB_04   , CAMPO_CAB_05   , CAMPO_CAB_06   , CAMPO_CAB_07 , CAMPO_CAB_08  , CAMPO_CAB_09,"
'   g_str_Parame = g_str_Parame & "           CAMPO_CAB_10, CAMPO_CAB_11  , CAMPO_CAB_12 , CAMPO_CAB_13 , CAMPO_CAB_14   , CAMPO_CAB_15   , CAMPO_CAB_16   , CAMPO_CAB_17 , CAMPO_CAB_18_1, CAMPO_CAB_18_2,"
'   g_str_Parame = g_str_Parame & "           CAMPO_CAB_19, CAMPO_CAB_20  , CAMPO_CAB_21 , CAMPO_CAB_22 , CAMPO_CAB_23   , CAMPO_CAB_24   , CAMPO_CAB_25   , CAMPO_CAB_26 , CAMPO_CAB_27  , CAMPO_DET_01,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET_02, CAMPO_DET_03  , CAMPO_DET_04 , "
'   g_str_Parame = g_str_Parame & "           MIN(CAMPO_DET_05) AS CAMPO_DET_05, "
'   g_str_Parame = g_str_Parame & "           CAMPO_DET_06, CAMPO_DET_07  , CAMPO_DET_08 , CAMPO_DET_09 , CAMPO_DET_10_1 , CAMPO_DET_10_2 , CAMPO_DET_10_3 , CAMPO_DET_11 , CAMPO_DET_12  , CAMPO_DET_13,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET_14, CAMPO_DET_15  , CAMPO_DET_16 , CAMPO_DET_17 , CAMPO_DET_18   , CAMPO_DET_19   , CAMPO_DET_20   , CAMPO_DET_21 , CAMPO_DET_22  , CAMPO_DET_23,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET_24, CAMPO_DET_25  , CAMPO_DET_26 , CAMPO_DET_27 , CAMPO_DET_28   , CAMPO_DET_29   , CAMPO_DET_30   , CAMPO_DET_31 , CAMPO_DET_32  , CAMPO_DET_33,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET_34, CAMPO_DET_35  , CAMPO_DET_36 , CAMPO_DET_37 , CAMPO_DET2_01  , CAMPO_DET2_02  , CAMPO_DET2_03  , CAMPO_DET2_04,"
'   g_str_Parame = g_str_Parame & "           MIN(CAMPO_DET2_05) AS CAMPO_DET2_05, "
'   g_str_Parame = g_str_Parame & "           CAMPO_DET2_06, CAMPO_DET2_07, CAMPO_DET2_08, CAMPO_DET2_09, CAMPO_DET2_10_1, CAMPO_DET2_10_2, CAMPO_DET2_10_3, CAMPO_DET2_11, CAMPO_DET2_12 , CAMPO_DET2_13,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET2_14, CAMPO_DET2_15, CAMPO_DET2_16, CAMPO_DET2_17, CAMPO_DET2_18  , CAMPO_DET2_19  , CAMPO_DET2_20  , CAMPO_DET2_21, CAMPO_DET2_22 , CAMPO_DET2_23,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET2_24, CAMPO_DET2_25, CAMPO_DET2_26, CAMPO_DET2_27, CAMPO_DET2_28  , CAMPO_DET2_29  , CAMPO_DET2_30  , CAMPO_DET2_31, CAMPO_DET2_32 , CAMPO_DET2_33,"
'   g_str_Parame = g_str_Parame & "           CAMPO_DET2_34, CAMPO_DET2_35, CAMPO_DET2_36, CAMPO_DET2_37, CAMPO_ADI_01   , CAMPO_ADI_02   , CAMPO_ADI_03   , CAMPO_ADI_04 , OPERACION     , NUMERO_MOVIMIENTO,"
'   g_str_Parame = g_str_Parame & "           FECHA_CANCELACION , SITUACION, FECHA_DEPOSITO "
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = "03" Then
         g_str_Parame = g_str_Parame & "    SELECT CAMPO_IDE_01,      CAMPO_IDE_02,       CAMPO_IDE_03,       CAMPO_IDE_04,       CAMPO_IDE_05,       CAMPO_IDE_06,      CAMPO_IDE_07,      CAMPO_IDE_08,      CAMPO_EMI_01,      CAMPO_EMI_02,"
         g_str_Parame = g_str_Parame & "           CAMPO_EMI_03,      CAMPO_EMI_04,       CAMPO_EMI_05,       CAMPO_EMI_06,       CAMPO_EMI_07,       CAMPO_EMI_08,      CAMPO_EMI_09,      CAMPO_EMI_10,      CAMPO_EMI_11,      CAMPO_EMI_12,"
         g_str_Parame = g_str_Parame & "           CAMPO_EMI_13,      CAMPO_EMI_14,       CAMPO_EMI_15,       CAMPO_EMI_16,       CAMPO_REC_01,       CAMPO_REC_02,      CAMPO_REC_03,      CAMPO_REC_04,      CAMPO_REC_05,      CAMPO_REC_06,"
         g_str_Parame = g_str_Parame & "           CAMPO_REC_07,      CAMPO_REC_08,       CAMPO_REC_09,       CAMPO_REC_10,       CAMPO_REC_11,       CAMPO_REC_12,      CAMPO_DRF_01,      CAMPO_DRF_02,      CAMPO_DRF_03,      CAMPO_DRF_04,"
         g_str_Parame = g_str_Parame & "           CAMPO_DRF_05,      CAMPO_DRF_06,       CAMPO_CAB_01,       CAMPO_CAB_02,       CAMPO_CAB_03,       CAMPO_CAB_04,      CAMPO_CAB_05,      CAMPO_CAB_06,      CAMPO_CAB_07,      CAMPO_CAB_08,"
         g_str_Parame = g_str_Parame & "           CAMPO_CAB_09,      CAMPO_CAB_10,       CAMPO_CAB_11,       CAMPO_CAB_12,       CAMPO_CAB_13,       CAMPO_CAB_14,      CAMPO_CAB_15,      CAMPO_CAB_16,      CAMPO_CAB_17,      CAMPO_CAB_18_1,"
         g_str_Parame = g_str_Parame & "           CAMPO_CAB_18_2,    CAMPO_CAB_19,       CAMPO_CAB_20,       CAMPO_CAB_21,       CAMPO_CAB_22,       CAMPO_CAB_23,      CAMPO_CAB_24,      CAMPO_CAB_25,      CAMPO_CAB_26,      CAMPO_CAB_27,"
         g_str_Parame = g_str_Parame & "           CAMPO_CAB_28_1,    CAMPO_CAB_28_2,     CAMPO_CAB_28_3,     CAMPO_CAB_28_4,     CAMPO_CAB_28_5,     CAMPO_DET_01,      CAMPO_DET_02,      CAMPO_DET_03,      CAMPO_DET_04,"
         g_str_Parame = g_str_Parame & "           MAX(CAMPO_DET_05) AS CAMPO_DET_05,"
         g_str_Parame = g_str_Parame & "           CAMPO_DET_06,      CAMPO_DET_07,       CAMPO_DET_08,       CAMPO_DET_09,       CAMPO_DET_10_1,     CAMPO_DET_10_2,    CAMPO_DET_10_3,    CAMPO_DET_10_4,    CAMPO_DET_10_5,    CAMPO_DET_11,"
         g_str_Parame = g_str_Parame & "           CAMPO_DET_12,      CAMPO_DET_13,       CAMPO_DET_14,       CAMPO_DET_15_1_1,   CAMPO_DET_15_2_1,   CAMPO_DET_15_1_2,  CAMPO_DET_15_2_2,  CAMPO_DET_15_1_3,  CAMPO_DET_15_2_3,  CAMPO_DET_15_1_4,"
         g_str_Parame = g_str_Parame & "           CAMPO_DET_15_2_4,  CAMPO_DET_15_1_5,   CAMPO_DET_15_2_5,   CAMPO_DET_15_1_6,   CAMPO_DET_15_2_6,   CAMPO_DET_15_1_7,  CAMPO_DET_15_2_7,  CAMPO_DET_15_1_8,  CAMPO_DET_15_2_8,  CAMPO_DET_15_1_9,"
         g_str_Parame = g_str_Parame & "           CAMPO_DET_15_2_9,  CAMPO_DET_15_1_10,  CAMPO_DET_15_2_10,  CAMPO_DET_15_1_11,  CAMPO_DET_15_2_11,  CAMPO_DET_37,      CAMPO_DET_38,      CAMPO_DET_39_1,    CAMPO_DET_39_2,    CAMPO_DET_39_3,"
         g_str_Parame = g_str_Parame & "           CAMPO_DET_39_4,    CAMPO_DET_39_5,     CAMPO_DET2_01,      CAMPO_DET2_02,      CAMPO_DET2_03,      CAMPO_DET2_04,"
         g_str_Parame = g_str_Parame & "           MAX(CAMPO_DET2_05) AS CAMPO_DET2_05,"
         g_str_Parame = g_str_Parame & "           CAMPO_DET2_06,     CAMPO_DET2_07,      CAMPO_DET2_08,      CAMPO_DET2_09,      CAMPO_DET2_10_1,    CAMPO_DET2_10_2,   CAMPO_DET2_10_3,   CAMPO_DET2_10_4,   CAMPO_DET2_10_5,   CAMPO_DET2_11,"
         g_str_Parame = g_str_Parame & "           CAMPO_DET2_12,     CAMPO_DET2_13,      CAMPO_DET2_14,      CAMPO_DET2_15_1_1,  CAMPO_DET2_15_2_1,  CAMPO_DET2_15_1_2, CAMPO_DET2_15_2_2, CAMPO_DET2_15_1_3, CAMPO_DET2_15_2_3, CAMPO_DET2_15_1_4,"
         g_str_Parame = g_str_Parame & "           CAMPO_DET2_15_2_4, CAMPO_DET2_15_1_5,  CAMPO_DET2_15_2_5,  CAMPO_DET2_15_1_6,  CAMPO_DET2_15_2_6,  CAMPO_DET2_15_1_7, CAMPO_DET2_15_2_7, CAMPO_DET2_15_1_8, CAMPO_DET2_15_2_8, CAMPO_DET2_15_1_9,"
         g_str_Parame = g_str_Parame & "           CAMPO_DET2_15_2_9, CAMPO_DET2_15_1_10, CAMPO_DET2_15_2_10, CAMPO_DET2_15_1_11, CAMPO_DET2_15_2_11, CAMPO_DET2_37,     CAMPO_DET2_38,     CAMPO_DET2_39_1,   CAMPO_DET2_39_2,   CAMPO_DET2_39_3,"
         g_str_Parame = g_str_Parame & "           CAMPO_DET2_39_4,   CAMPO_DET2_39_5,    CAMPO_ADI_01,       CAMPO_ADI_02,       CAMPO_ADI_03,       CAMPO_ADI_04,      OPERACION,         NUMERO_MOVIMIENTO, FECHA_CANCELACION, SITUACION,"
         g_str_Parame = g_str_Parame & "           FECHA_DEPOSITO"
         g_str_Parame = g_str_Parame & "     FROM ( "
         g_str_Parame = g_str_Parame & "              SELECT 'IDE'                                                                                               AS CAMPO_IDE_01, "
         g_str_Parame = g_str_Parame & "                     'B'                                                                                                 AS CAMPO_IDE_02, "
         g_str_Parame = g_str_Parame & "                     SUBSTR(CAJMOV_FECDEP,1,4) || '-' || SUBSTR(CAJMOV_FECDEP,5,2) || '-' || SUBSTR(CAJMOV_FECDEP,7,2)   AS CAMPO_IDE_03, "
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_IDE_04,"
         g_str_Parame = g_str_Parame & "                     '03'                                                                                                AS CAMPO_IDE_05,              " '--CATALOGO N°1.- FACTURA
         g_str_Parame = g_str_Parame & "                     C.CATSUN_CODIGO                                                                                     AS CAMPO_IDE_06,              " '--CATALOGO N°2.-
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_IDE_07,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_IDE_08,"
         
         g_str_Parame = g_str_Parame & "                     'EMI'                                                                                               AS CAMPO_EMI_01,"
         g_str_Parame = g_str_Parame & "                     'B'                                                                                                 AS CAMPO_EMI_02,"
         g_str_Parame = g_str_Parame & "                     '6'                                                                                                 AS CAMPO_EMI_03,              " '--CATALOGO N°6.-
         g_str_Parame = g_str_Parame & "                     '20511904162'                                                                                       AS CAMPO_EMI_04,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_05,"
         g_str_Parame = g_str_Parame & "                     'EDPYME MICASITA SA'                                                                                AS CAMPO_EMI_06,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_07,              " '--ATALOGO N°13.-UBIGEO
         g_str_Parame = g_str_Parame & "                     'AV RIVERA NAVARRETE 645'                                                                           AS CAMPO_EMI_08,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_09,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_10,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_11,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_12,"
         g_str_Parame = g_str_Parame & "                     'PE'                                                                                                AS CAMPO_EMI_13,              " '--CATALOGO N°4.-
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_14,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_EMI_15,"
         g_str_Parame = g_str_Parame & "                     '0000'                                                                                              AS CAMPO_EMI_16,"
         
         g_str_Parame = g_str_Parame & "                     'REC'                                                                                               AS CAMPO_REC_01,"
         g_str_Parame = g_str_Parame & "                     'B'                                                                                                 AS CAMPO_REC_02,"
         g_str_Parame = g_str_Parame & "                     TRIM(A.CAJMOV_TIPDOC)                                                                               AS CAMPO_REC_03,              " '--CATALOGO N°6.-
         g_str_Parame = g_str_Parame & "                     TRIM(A.CAJMOV_NUMDOC)                                                                               AS CAMPO_REC_04,"
         g_str_Parame = g_str_Parame & "                     TRIM(D.DATGEN_APEPAT) || ' ' || TRIM(D.DATGEN_APEMAT) || ' ' || TRIM(D.DATGEN_NOMBRE)               AS CAMPO_REC_05,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_REC_06,"
         g_str_Parame = g_str_Parame & "                     TRIM(H.PARDES_DESCRI)                                                                               AS CAMPO_REC_07,"
         g_str_Parame = g_str_Parame & "                     TRIM(I.PARDES_DESCRI)                                                                               AS CAMPO_REC_08,"
         g_str_Parame = g_str_Parame & "                     TRIM(G.PARDES_DESCRI)                                                                               AS CAMPO_REC_09,"
         g_str_Parame = g_str_Parame & "                     'PE'                                                                                                AS CAMPO_REC_10,"
         g_str_Parame = g_str_Parame & "                     TRIM(D.DATGEN_TELEFO)                                                                               AS CAMPO_REC_11,"
         g_str_Parame = g_str_Parame & "                     TRIM(D.DATGEN_DIRELE)                                                                               AS CAMPO_REC_12,"
         
         g_str_Parame = g_str_Parame & "                     'DRF'                                                                                               AS CAMPO_DRF_01,"
         g_str_Parame = g_str_Parame & "                     'B'                                                                                                 AS CAMPO_DRF_02,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DRF_03,               "
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DRF_04,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DRF_05,               " '--PARA NOTA CREDITO/DEBITO
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DRF_06,"
         
         g_str_Parame = g_str_Parame & "                     'CAB'                                                                                               AS CAMPO_CAB_01,"
         g_str_Parame = g_str_Parame & "                     'B'                                                                                                 AS CAMPO_CAB_02,"
         g_str_Parame = g_str_Parame & "                     '1001'                                                                                              AS CAMPO_CAB_03,               " '--CATALOGO N°14.-
         g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_CAB_04,"
         g_str_Parame = g_str_Parame & "                     '1002'                                                                                              AS CAMPO_CAB_05,               " '--CATALOGO N°14.-
         g_str_Parame = g_str_Parame & "                     A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_06,"
         g_str_Parame = g_str_Parame & "                     '1003'                                                                                              AS CAMPO_CAB_07,               " '--CATALOGO N°14.-
         g_str_Parame = g_str_Parame & "                     '0.00'                                                                                              AS CAMPO_CAB_08,"
         g_str_Parame = g_str_Parame & "                     '1004'                                                                                              AS CAMPO_CAB_09,               " '--CATALOGO N°14.-
         g_str_Parame = g_str_Parame & "                     '0.00'                                                                                              AS CAMPO_CAB_10,"
         g_str_Parame = g_str_Parame & "                     '1000'                                                                                              AS CAMPO_CAB_11,               " '--CATALOGO N°14.-
         g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_CAB_12,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_13,"
         
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_14,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_15,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_16,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_17,"
         g_str_Parame = g_str_Parame & "                     '9998'                                                                                              AS CAMPO_CAB_18_1,"
         g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_CAB_18_2,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_19,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_20,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_21,"
         g_str_Parame = g_str_Parame & "                     A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_22,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_23,                " '23.1 , 23.2 , 23.3 , 23.4 , 23.5 , 23.6 , 23.7 , 23.8
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_24,"
         g_str_Parame = g_str_Parame & "                     '0112'                                                                                              AS CAMPO_CAB_25,                " '                 --CATALOGO N°17.-
         g_str_Parame = g_str_Parame & "                     '[1000'                                                                                             AS CAMPO_CAB_26,                " '26.1 , 26.2      --CATALOGO N°15.- DETALLE EN LETRAS DEL IMPORTE
         g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_CAB_27,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_28_1,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_28_2,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_28_3,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_28_4,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_CAB_28_5,"
         
         g_str_Parame = g_str_Parame & "                     'DET1'                                                                                              AS CAMPO_DET_01,"
         g_str_Parame = g_str_Parame & "                     'B'                                                                                                 AS CAMPO_DET_02,"
         g_str_Parame = g_str_Parame & "                     '001'                                                                                               AS CAMPO_DET_03,                " '-- Número de orden de ítem
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_04,                "
         g_str_Parame = g_str_Parame & "                     'INTERES ' ||"
         g_str_Parame = g_str_Parame & "                     SUBSTR(TO_CHAR(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(HIPCUO_FECVCT,1,4)  AS CAMPO_DET_05," 'CAJMOV_FECDEP
         g_str_Parame = g_str_Parame & "                     1.000                                                                                               AS CAMPO_DET_06,"
         g_str_Parame = g_str_Parame & "                     'NIU'                                                                                               AS CAMPO_DET_07,                " '--CATALOGO N°3.-
         g_str_Parame = g_str_Parame & "                     A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_08,"
         g_str_Parame = g_str_Parame & "                     1.000 * A.CAJMOV_INTERE                                                                             AS CAMPO_DET_09,"
         g_str_Parame = g_str_Parame & "                     '9998'                                                                                              AS CAMPO_DET_10_1,              " '--CATALOGO N°5,7 u 8.-
         g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_DET_10_2,"
         g_str_Parame = g_str_Parame & "                     '30'                                                                                                AS CAMPO_DET_10_3,"
         g_str_Parame = g_str_Parame & "                     1.000 * A.CAJMOV_INTERE                                                                             AS CAMPO_DET_10_4,"
         g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_DET_10_5,"
         g_str_Parame = g_str_Parame & "                     A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_11,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_12,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_13,"
         g_str_Parame = g_str_Parame & "                     '84121901'                                                                                          AS CAMPO_DET_14,                " '--CATALOGO N°15.-
         g_str_Parame = g_str_Parame & "                     '7001'                                                                                              AS CAMPO_DET_15_1_1,"
         g_str_Parame = g_str_Parame & "                     '1'                                                                                                 AS CAMPO_DET_15_2_1,                " '--CATALOGO N°26.- Si es construcción /adquisición
         g_str_Parame = g_str_Parame & "                     '7002'                                                                                              AS CAMPO_DET_15_1_2,"
         g_str_Parame = g_str_Parame & "                     CASE WHEN L.SOLMAE_PRIVIV = 1 THEN 3  "
         g_str_Parame = g_str_Parame & "                          WHEN L.SOLMAE_PRIVIV = 2 THEN 0 END                                                            AS CAMPO_DET_15_2_2,                " '--CATALOGO N°27.- VERIFICAR
         g_str_Parame = g_str_Parame & "                     '7003'                                                                                              AS CAMPO_DET_15_1_3,"
         g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_3,"
         g_str_Parame = g_str_Parame & "                     '7004'                                                                                              AS CAMPO_DET_15_1_4,"
         g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_4,                " '--NRO CONTRATO O PRESTAMO
         g_str_Parame = g_str_Parame & "                     '7005'                                                                                              AS CAMPO_DET_15_1_5,"
         g_str_Parame = g_str_Parame & "                     SUBSTR(J.HIPMAE_FECACT,1,4) || '-' || SUBSTR(J.HIPMAE_FECACT,5,2) || '-' || SUBSTR(J.HIPMAE_FECACT,7,2)"
         g_str_Parame = g_str_Parame & "                                                                                                                         AS CAMPO_DET_15_2_5,"
         g_str_Parame = g_str_Parame & "                     '7006'                                                                                              AS CAMPO_DET_15_1_6,"
         g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_6,"                 '/*P.SOLINM_UBIGEO*/
         g_str_Parame = g_str_Parame & "                     '7007'                                                                                              AS CAMPO_DET_15_1_7,"
         g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_7,"
         g_str_Parame = g_str_Parame & "                     '7008'                                                                                              AS CAMPO_DET_15_1_8,"
         g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_8,"
         g_str_Parame = g_str_Parame & "                     '7009'                                                                                              AS CAMPO_DET_15_1_9,"
         g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_9,"
         g_str_Parame = g_str_Parame & "                     '7010'                                                                                              AS CAMPO_DET_15_1_10,"
         g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_10,"
         g_str_Parame = g_str_Parame & "                     '7011'                                                                                              AS CAMPO_DET_15_1_11,"
         g_str_Parame = g_str_Parame & "                     '-'                                                                                                 AS CAMPO_DET_15_2_11,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_37,"
         g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_DET_38,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_39_1,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_39_2,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_39_3,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_39_4,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET_39_5,"
         
         g_str_Parame = g_str_Parame & "                     'DET2'                                                                                              AS CAMPO_DET2_01,"
         g_str_Parame = g_str_Parame & "                     'B'                                                                                                 AS CAMPO_DET2_02,"
         g_str_Parame = g_str_Parame & "                     '002'                                                                                               AS CAMPO_DET2_03,                " '-- Número de orden de ítem
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_04,                "
         g_str_Parame = g_str_Parame & "                     'OTROS IMPORTES ' ||"
         g_str_Parame = g_str_Parame & "                     SUBSTR(TO_CHAR(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(HIPCUO_FECVCT,1,4)  AS CAMPO_DET2_05," 'CAJMOV_FECDEP
         g_str_Parame = g_str_Parame & "                     1.000                                                                                               AS CAMPO_DET2_06,"
         g_str_Parame = g_str_Parame & "                     'NIU'                                                                                               AS CAMPO_DET2_07,                " '--CATALOGO N°3.-
         g_str_Parame = g_str_Parame & "                     (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)                                                                 AS CAMPO_DET2_08,"
         g_str_Parame = g_str_Parame & "                     1.000 * (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)                                                         AS CAMPO_DET2_09,"
         g_str_Parame = g_str_Parame & "                     '9998'                                                                                              AS CAMPO_DET2_10_1,                 " '--CATALOGO N°5,7 u 8.-
         g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_DET2_10_2,"
         g_str_Parame = g_str_Parame & "                     '30'                                                                                                AS CAMPO_DET2_10_3,"
         g_str_Parame = g_str_Parame & "                     1.000 * A.CAJMOV_INTERE                                                                             AS CAMPO_DET2_10_4,"
         g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_DET2_10_5,"
         g_str_Parame = g_str_Parame & "                     A.CAJMOV_IMPPAG - A.CAJMOV_INTERE                                                                   AS CAMPO_DET2_11,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_12,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_13,"
         g_str_Parame = g_str_Parame & "                     '84121501'                                                                                          AS CAMPO_DET2_14,                    " '--CATALOGO N°15.-
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_1,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_1,                " '--NRO CONTRATO O PRESTAMO
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_2,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_2,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_3,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_3,                " '--CATALOGO N°26.- Si es construcción /adquisición
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_4,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_4,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_5,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_5,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_6,                " '--CATALOGO N°27.- VERIFICAR
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_6,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_7,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_7,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_8,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_8,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_9,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_9,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_10,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_10,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_1_11,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_15_2_11,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_37,"
         g_str_Parame = g_str_Parame & "                     0.00                                                                                                AS CAMPO_DET2_38,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_39_1,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_39_2,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_39_3,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_39_4,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_DET2_39_5,"
      
         g_str_Parame = g_str_Parame & "                     'ADI1'                                                                                              AS CAMPO_ADI_01,"
         g_str_Parame = g_str_Parame & "                     'B'                                                                                                 AS CAMPO_ADI_02,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_ADI_03,"
         g_str_Parame = g_str_Parame & "                     ''                                                                                                  AS CAMPO_ADI_04,"
         g_str_Parame = g_str_Parame & "                     TRIM(A.CAJMOV_NUMOPE)                                                                               AS OPERACION, "
         g_str_Parame = g_str_Parame & "                     CAJMOV_NUMMOV                                                                                       AS NUMERO_MOVIMIENTO , "
         g_str_Parame = g_str_Parame & "                     HIPMAE_FECCAN                                                                                       AS FECHA_CANCELACION ,"
         g_str_Parame = g_str_Parame & "                     HIPMAE_SITUAC                                                                                       AS SITUACION, "
         g_str_Parame = g_str_Parame & "                     CAJMOV_FECDEP                                                                                       AS FECHA_DEPOSITO "
         g_str_Parame = g_str_Parame & "                FROM OPE_CAJMOV A "
         g_str_Parame = g_str_Parame & "                     INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '204' AND B.PARDES_CODITE = A.CAJMOV_MONPAG "
         g_str_Parame = g_str_Parame & "                     INNER JOIN CNTBL_CATSUN C ON C.CATSUN_NROCAT = 2 AND C.CATSUN_DESCRI = TRIM(B.PARDES_DESCRI) "
         g_str_Parame = g_str_Parame & "                     INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = A.CAJMOV_TIPDOC AND D.DATGEN_NUMDOC = A.CAJMOV_NUMDOC "
         g_str_Parame = g_str_Parame & "                     INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 201 AND E.PARDES_CODITE = D.DATGEN_TIPVIA "
         g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 202 AND F.PARDES_CODITE = D.DATGEN_TIPZON "
         g_str_Parame = g_str_Parame & "                     INNER JOIN MNT_PARDES G ON G.PARDES_CODGRP = 101 AND G.PARDES_CODITE = D.DATGEN_UBIGEO "
         g_str_Parame = g_str_Parame & "                     INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 101 AND H.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,2)||'0000' "
         g_str_Parame = g_str_Parame & "                     INNER JOIN MNT_PARDES I ON I.PARDES_CODGRP = 101 AND I.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,4)||'00' "
         g_str_Parame = g_str_Parame & "                     INNER JOIN CRE_HIPMAE J ON J.HIPMAE_NUMOPE = A.CAJMOV_NUMOPE "
         g_str_Parame = g_str_Parame & "                      LEFT JOIN CRE_HIPGAR K ON K.HIPGAR_NUMOPE = J.HIPMAE_NUMOPE AND K.HIPGAR_BIEGAR = 1 "
         g_str_Parame = g_str_Parame & "                      LEFT JOIN CRE_SOLMAE L ON L.SOLMAE_NUMERO = J.HIPMAE_NUMSOL "
         g_str_Parame = g_str_Parame & "                      LEFT JOIN CRE_SOLINM P ON P.SOLINM_NUMSOL = J.HIPMAE_NUMSOL "
         g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES Q ON Q.PARDES_CODGRP = 201 AND Q.PARDES_CODITE = P.SOLINM_TIPVIA "
         g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES R ON R.PARDES_CODGRP = 202 AND R.PARDES_CODITE = P.SOLINM_TIPZON "
         g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES S ON S.PARDES_CODGRP = 101 AND S.PARDES_CODITE = P.SOLINM_UBIGEO "
         g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES T ON T.PARDES_CODGRP = 101 AND T.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,4)||'00' "
         g_str_Parame = g_str_Parame & "                      LEFT JOIN MNT_PARDES U ON U.PARDES_CODGRP = 101 AND U.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,2)||'0000' "
         g_str_Parame = g_str_Parame & "                     INNER JOIN CRE_HIPPAG V ON V.HIPPAG_NUMOPE = A.CAJMOV_NUMOPE AND V.HIPPAG_FECPAG = A.CAJMOV_FECDEP AND V.HIPPAG_NUMMOV = A.CAJMOV_NUMMOV"
         g_str_Parame = g_str_Parame & "                     INNER JOIN CRE_HIPCUO W ON W.HIPCUO_NUMOPE = V.HIPPAG_NUMOPE AND W.HIPCUO_TIPCRO = 1 AND W.HIPCUO_NUMCUO = V.HIPPAG_NUMCUO"
         g_str_Parame = g_str_Parame & "               WHERE CAJMOV_SUCMOV IS NOT NULL "
         g_str_Parame = g_str_Parame & "                 AND CAJMOV_USUMOV IS NOT NULL "
         g_str_Parame = g_str_Parame & "                 AND CAJMOV_FECMOV > 0"
         g_str_Parame = g_str_Parame & "                 AND CAJMOV_NUMMOV > 0 "
         g_str_Parame = g_str_Parame & "                 AND CAJMOV_CODBAN IS NOT NULL "
         
         If Chk_FecAct.Value = 0 Then
            g_str_Parame = g_str_Parame & "              AND CAJMOV_FECDEP = '" & l_str_FecCar & "' "
         Else
            g_str_Parame = g_str_Parame & "              AND CAJMOV_FECDEP <= '" & l_str_FecCar & "' "
         End If
         
         g_str_Parame = g_str_Parame & "                 AND CAJMOV_FLGPRO = 0 "
         g_str_Parame = g_str_Parame & "                 AND CAJMOV_TIPMOV = '1102' "
         g_str_Parame = g_str_Parame & "                 AND CAJMOV_TIPDOC IN (1, 6) "                      '1-DNI y 6-RUC
         g_str_Parame = g_str_Parame & "               ORDER BY A.CAJMOV_FECMOV , A.CAJMOV_NUMMOV "
         g_str_Parame = g_str_Parame & "      )"
         g_str_Parame = g_str_Parame & "      GROUP BY  CAMPO_IDE_01,      CAMPO_IDE_02,       CAMPO_IDE_03,       CAMPO_IDE_04,       CAMPO_IDE_05,       CAMPO_IDE_06,      CAMPO_IDE_07,      CAMPO_IDE_08,      CAMPO_EMI_01,      CAMPO_EMI_02,"
         g_str_Parame = g_str_Parame & "                CAMPO_EMI_03,      CAMPO_EMI_04,       CAMPO_EMI_05,       CAMPO_EMI_06,       CAMPO_EMI_07,       CAMPO_EMI_08,      CAMPO_EMI_09,      CAMPO_EMI_10,      CAMPO_EMI_11,      CAMPO_EMI_12,"
         g_str_Parame = g_str_Parame & "                CAMPO_EMI_13,      CAMPO_EMI_14,       CAMPO_EMI_15,       CAMPO_EMI_16,       CAMPO_REC_01,       CAMPO_REC_02,      CAMPO_REC_03,      CAMPO_REC_04,      CAMPO_REC_05,      CAMPO_REC_06,"
         g_str_Parame = g_str_Parame & "                CAMPO_REC_07,      CAMPO_REC_08,       CAMPO_REC_09,       CAMPO_REC_10,       CAMPO_REC_11,       CAMPO_REC_12,      CAMPO_DRF_01,      CAMPO_DRF_02,      CAMPO_DRF_03,      CAMPO_DRF_04,"
         g_str_Parame = g_str_Parame & "                CAMPO_DRF_05,      CAMPO_DRF_06,       CAMPO_CAB_01,       CAMPO_CAB_02,       CAMPO_CAB_03,       CAMPO_CAB_04,      CAMPO_CAB_05,      CAMPO_CAB_06,      CAMPO_CAB_07,      CAMPO_CAB_08,"
         g_str_Parame = g_str_Parame & "                CAMPO_CAB_09,      CAMPO_CAB_10,       CAMPO_CAB_11,       CAMPO_CAB_12,       CAMPO_CAB_13,       CAMPO_CAB_14,      CAMPO_CAB_15,      CAMPO_CAB_16,      CAMPO_CAB_17,      CAMPO_CAB_18_1,"
         g_str_Parame = g_str_Parame & "                CAMPO_CAB_18_2,    CAMPO_CAB_19,       CAMPO_CAB_20,       CAMPO_CAB_21,       CAMPO_CAB_22,       CAMPO_CAB_23,      CAMPO_CAB_24,      CAMPO_CAB_25,      CAMPO_CAB_26,      CAMPO_CAB_27,"
         g_str_Parame = g_str_Parame & "                CAMPO_CAB_28_1,    CAMPO_CAB_28_2,     CAMPO_CAB_28_3,     CAMPO_CAB_28_4,     CAMPO_CAB_28_5,     CAMPO_DET_01,      CAMPO_DET_02,      CAMPO_DET_03,      CAMPO_DET_04,"
         '                                                CAMPO_DET_05,
         g_str_Parame = g_str_Parame & "                CAMPO_DET_06,      CAMPO_DET_07,       CAMPO_DET_08,       CAMPO_DET_09,       CAMPO_DET_10_1,     CAMPO_DET_10_2,    CAMPO_DET_10_3,    CAMPO_DET_10_4,    CAMPO_DET_10_5,    CAMPO_DET_11,"
         g_str_Parame = g_str_Parame & "                CAMPO_DET_12,      CAMPO_DET_13,       CAMPO_DET_14,       CAMPO_DET_15_1_1,   CAMPO_DET_15_2_1,   CAMPO_DET_15_1_2,  CAMPO_DET_15_2_2,  CAMPO_DET_15_1_3,  CAMPO_DET_15_2_3,  CAMPO_DET_15_1_4,"
         g_str_Parame = g_str_Parame & "                CAMPO_DET_15_2_4,  CAMPO_DET_15_1_5,   CAMPO_DET_15_2_5,   CAMPO_DET_15_1_6,   CAMPO_DET_15_2_6,   CAMPO_DET_15_1_7,  CAMPO_DET_15_2_7,  CAMPO_DET_15_1_8,  CAMPO_DET_15_2_8,  CAMPO_DET_15_1_9,"
         g_str_Parame = g_str_Parame & "                CAMPO_DET_15_2_9,  CAMPO_DET_15_1_10,  CAMPO_DET_15_2_10,  CAMPO_DET_15_1_11,  CAMPO_DET_15_2_11,  CAMPO_DET_37,      CAMPO_DET_38,      CAMPO_DET_39_1,    CAMPO_DET_39_2,    CAMPO_DET_39_3,"
         g_str_Parame = g_str_Parame & "                CAMPO_DET_39_4,    CAMPO_DET_39_5,     CAMPO_DET2_01,      CAMPO_DET2_02,      CAMPO_DET2_03,      CAMPO_DET2_04,"
         '                                                CAMPO_DET2_05,
         g_str_Parame = g_str_Parame & "                CAMPO_DET2_06,     CAMPO_DET2_07,      CAMPO_DET2_08,      CAMPO_DET2_09,      CAMPO_DET2_10_1,    CAMPO_DET2_10_2,   CAMPO_DET2_10_3,   CAMPO_DET2_10_4,   CAMPO_DET2_10_5,   CAMPO_DET2_11,"
         g_str_Parame = g_str_Parame & "                CAMPO_DET2_12,     CAMPO_DET2_13,      CAMPO_DET2_14,      CAMPO_DET2_15_1_1,  CAMPO_DET2_15_2_1,  CAMPO_DET2_15_1_2, CAMPO_DET2_15_2_2, CAMPO_DET2_15_1_3, CAMPO_DET2_15_2_3, CAMPO_DET2_15_1_4,"
         g_str_Parame = g_str_Parame & "                CAMPO_DET2_15_2_4, CAMPO_DET2_15_1_5,  CAMPO_DET2_15_2_5,  CAMPO_DET2_15_1_6,  CAMPO_DET2_15_2_6,  CAMPO_DET2_15_1_7, CAMPO_DET2_15_2_7, CAMPO_DET2_15_1_8, CAMPO_DET2_15_2_8, CAMPO_DET2_15_1_9,"
         g_str_Parame = g_str_Parame & "                CAMPO_DET2_15_2_9, CAMPO_DET2_15_1_10, CAMPO_DET2_15_2_10, CAMPO_DET2_15_1_11, CAMPO_DET2_15_2_11, CAMPO_DET2_37,     CAMPO_DET2_38,     CAMPO_DET2_39_1,   CAMPO_DET2_39_2,   CAMPO_DET2_39_3,"
         g_str_Parame = g_str_Parame & "                CAMPO_DET2_39_4,   CAMPO_DET2_39_5,    CAMPO_ADI_01,       CAMPO_ADI_02,       CAMPO_ADI_03,       CAMPO_ADI_04,      OPERACION,         NUMERO_MOVIMIENTO, FECHA_CANCELACION, SITUACION,"
         g_str_Parame = g_str_Parame & "                FECHA_DEPOSITO"
         
      
      '   g_str_Parame = g_str_Parame & "      GROUP BY CAMPO_IDE_01 , CAMPO_IDE_02  , CAMPO_IDE_03 , CAMPO_IDE_04 , CAMPO_IDE_05   , CAMPO_IDE_06   , CAMPO_IDE_07   , CAMPO_IDE_08  , CAMPO_EMI_01  , CAMPO_EMI_02,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_EMI_03 , CAMPO_EMI_04  , CAMPO_EMI_05 , CAMPO_EMI_06 , CAMPO_EMI_07   , CAMPO_EMI_08   , CAMPO_EMI_09   , CAMPO_EMI_10  , CAMPO_EMI_11  , CAMPO_EMI_12,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_EMI_13 , CAMPO_EMI_14  , CAMPO_EMI_15 , CAMPO_REC_01 , CAMPO_REC_02   , CAMPO_REC_03   , CAMPO_REC_04   , CAMPO_REC_05  , CAMPO_REC_06  , CAMPO_REC_07,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_REC_08 , CAMPO_REC_09  , CAMPO_REC_10 , CAMPO_REC_11 , CAMPO_REC_12   , CAMPO_DRF_01   , CAMPO_DRF_02   , CAMPO_DRF_03  , CAMPO_DRF_04  , CAMPO_DRF_05,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_DRF_06 , CAMPO_CAB_01  , CAMPO_CAB_02 , CAMPO_CAB_03 , CAMPO_CAB_04   , CAMPO_CAB_05   , CAMPO_CAB_06   , CAMPO_CAB_07  , CAMPO_CAB_08  , CAMPO_CAB_09,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_CAB_10 , CAMPO_CAB_11  , CAMPO_CAB_12 , CAMPO_CAB_13 , CAMPO_CAB_14   , CAMPO_CAB_15   , CAMPO_CAB_16   , CAMPO_CAB_17  , CAMPO_CAB_18_1, CAMPO_CAB_18_2,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_CAB_19 , CAMPO_CAB_20  , CAMPO_CAB_21 , CAMPO_CAB_22 , CAMPO_CAB_23   , CAMPO_CAB_24   , CAMPO_CAB_25   , CAMPO_CAB_26  , CAMPO_CAB_27  , CAMPO_DET_01,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_DET_02 , CAMPO_DET_03  , CAMPO_DET_04 , "
      '   g_str_Parame = g_str_Parame & "               CAMPO_DET_06 , CAMPO_DET_07  , CAMPO_DET_08 , CAMPO_DET_09 , CAMPO_DET_10_1 , CAMPO_DET_10_2 , CAMPO_DET_10_3 , CAMPO_DET_11  , CAMPO_DET_12  , CAMPO_DET_13,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_DET_14 , CAMPO_DET_15  , CAMPO_DET_16 , CAMPO_DET_17 , CAMPO_DET_18   , CAMPO_DET_19   , CAMPO_DET_20   , CAMPO_DET_21  , CAMPO_DET_22  , CAMPO_DET_23,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_DET_24 , CAMPO_DET_25  , CAMPO_DET_26 , CAMPO_DET_27 , CAMPO_DET_28   , CAMPO_DET_29   , CAMPO_DET_30   , CAMPO_DET_31  , CAMPO_DET_32  , CAMPO_DET_33,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_DET_34 , CAMPO_DET_35  , CAMPO_DET_36 , CAMPO_DET_37 , CAMPO_DET2_01  , CAMPO_DET2_02  , CAMPO_DET2_03  , CAMPO_DET2_04 , "
      '   g_str_Parame = g_str_Parame & "               CAMPO_DET2_06, CAMPO_DET2_07, CAMPO_DET2_08 , CAMPO_DET2_09, CAMPO_DET2_10_1, CAMPO_DET2_10_2, CAMPO_DET2_10_3, CAMPO_DET2_11 , CAMPO_DET2_12 , CAMPO_DET2_13,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_DET2_14, CAMPO_DET2_15, CAMPO_DET2_16 , CAMPO_DET2_17, CAMPO_DET2_18  , CAMPO_DET2_19  , CAMPO_DET2_20  , CAMPO_DET2_21 , CAMPO_DET2_22 , CAMPO_DET2_23,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_DET2_24, CAMPO_DET2_25, CAMPO_DET2_26 , CAMPO_DET2_27, CAMPO_DET2_28  , CAMPO_DET2_29  , CAMPO_DET2_30  , CAMPO_DET2_31 , CAMPO_DET2_32 , CAMPO_DET2_33,"
      '   g_str_Parame = g_str_Parame & "               CAMPO_DET2_34, CAMPO_DET2_35, CAMPO_DET2_36 , CAMPO_DET2_37, CAMPO_ADI_01   , CAMPO_ADI_02   , CAMPO_ADI_03   , CAMPO_ADI_04  , OPERACION     , NUMERO_MOVIMIENTO,"
      '   g_str_Parame = g_str_Parame & "               FECHA_CANCELACION , SITUACION, FECHA_DEPOSITO"
   
   ElseIf cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = "05" Then
      
      g_str_Parame = g_str_Parame & "        SELECT 'IDE'                                                                                               AS CAMPO_IDE_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_IDE_02, "
      g_str_Parame = g_str_Parame & "               SUBSTR(PPGCAB_FECPPG,1,4) || '-' || SUBSTR(PPGCAB_FECPPG,5,2) || '-' || SUBSTR(PPGCAB_FECPPG,7,2)   AS CAMPO_IDE_03, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_IDE_04, "
      g_str_Parame = g_str_Parame & "               '03'                                                                                                AS CAMPO_IDE_05, "
      g_str_Parame = g_str_Parame & "               C.CATSUN_CODIGO                                                                                     AS CAMPO_IDE_06, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_IDE_07, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_IDE_08, "
      
      g_str_Parame = g_str_Parame & "               'EMI'                                                                                               AS CAMPO_EMI_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_EMI_02, "
      g_str_Parame = g_str_Parame & "               '6'                                                                                                 AS CAMPO_EMI_03, "
      g_str_Parame = g_str_Parame & "               '20511904162'                                                                                       AS CAMPO_EMI_04, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_05, "
      g_str_Parame = g_str_Parame & "               'EDPYME MICASITA SA'                                                                                AS CAMPO_EMI_06, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_07, "
      g_str_Parame = g_str_Parame & "               'AV RIVERA NAVARRETE 645'                                                                           AS CAMPO_EMI_08, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_09, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_10, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_11, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_12, "
      g_str_Parame = g_str_Parame & "               'PE'                                                                                                AS CAMPO_EMI_13, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_14, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_EMI_15, "
      g_str_Parame = g_str_Parame & "               '0000'                                                                                              AS CAMPO_EMI_16, "
           
      g_str_Parame = g_str_Parame & "               'REC'                                                                                               AS CAMPO_REC_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_REC_02, "
      g_str_Parame = g_str_Parame & "               TRIM(J.HIPMAE_TDOCLI)                                                                               AS CAMPO_REC_03, "
      g_str_Parame = g_str_Parame & "               TRIM(J.HIPMAE_NDOCLI)                                                                               AS CAMPO_REC_04, "
      g_str_Parame = g_str_Parame & "               TRIM(D.DATGEN_APEPAT) || ' ' || TRIM(D.DATGEN_APEMAT) || ' ' || TRIM(D.DATGEN_NOMBRE)               AS CAMPO_REC_05, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_REC_06, "
      g_str_Parame = g_str_Parame & "               TRIM(H.PARDES_DESCRI)                                                                               AS CAMPO_REC_07, "
      g_str_Parame = g_str_Parame & "               TRIM(I.PARDES_DESCRI)                                                                               AS CAMPO_REC_08, "
      g_str_Parame = g_str_Parame & "               TRIM(G.PARDES_DESCRI)                                                                               AS CAMPO_REC_09, "
      g_str_Parame = g_str_Parame & "               'PE'                                                                                                AS CAMPO_REC_10, "
      g_str_Parame = g_str_Parame & "               TRIM(D.DATGEN_TELEFO)                                                                               AS CAMPO_REC_11, "
      g_str_Parame = g_str_Parame & "               TRIM(D.DATGEN_DIRELE)                                                                               AS CAMPO_REC_12, "
      
      g_str_Parame = g_str_Parame & "               'DRF'                                                                                               AS CAMPO_DRF_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_DRF_02, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DRF_03, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DRF_04, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DRF_05, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DRF_06, "
      
      g_str_Parame = g_str_Parame & "               'CAB'                                                                                               AS CAMPO_CAB_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_CAB_02, "
      g_str_Parame = g_str_Parame & "               '1001'                                                                                              AS CAMPO_CAB_03, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_CAB_04, "
      g_str_Parame = g_str_Parame & "               '1002'                                                                                              AS CAMPO_CAB_05, "
      g_str_Parame = g_str_Parame & "               CASE WHEN A.PPGCAB_TIPPPG = 1 THEN A.PPGCAB_MTODEP ELSE PPGCAB_MTOTOT END                           AS CAMPO_CAB_06, "
      g_str_Parame = g_str_Parame & "               '1003'                                                                                              AS CAMPO_CAB_07, "
      g_str_Parame = g_str_Parame & "               '0.00'                                                                                              AS CAMPO_CAB_08, "
      g_str_Parame = g_str_Parame & "               '1004'                                                                                              AS CAMPO_CAB_09, "
      g_str_Parame = g_str_Parame & "               '0.00'                                                                                              AS CAMPO_CAB_10, "
      g_str_Parame = g_str_Parame & "               '1000'                                                                                              AS CAMPO_CAB_11, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_CAB_12, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_13, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_14, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_15, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_16, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_17, "
      g_str_Parame = g_str_Parame & "               '9998'                                                                                              AS CAMPO_CAB_18_1, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_CAB_18_2, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_19, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_20, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_21, "
      g_str_Parame = g_str_Parame & "               CASE WHEN A.PPGCAB_TIPPPG = 1 THEN A.PPGCAB_MTODEP ELSE PPGCAB_MTOTOT END                           AS CAMPO_CAB_22, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_23, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_24, "
      g_str_Parame = g_str_Parame & "               '0112'                                                                                              AS CAMPO_CAB_25, "
      'g_str_Parame = g_str_Parame & "               '13'                                                                                                AS CAMPO_CAB_26, "
      g_str_Parame = g_str_Parame & "               '[1000'                                                                                             AS CAMPO_CAB_26, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_CAB_27, "
      g_str_Parame = g_str_Parame & "               'false'                                                                                             AS CAMPO_CAB_28_1,"
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_28_2,"
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_28_3,"
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_28_4,"
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_CAB_28_5,"
      
      g_str_Parame = g_str_Parame & "               'DET1'                                                                                              AS CAMPO_DET_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_DET_02, "
      g_str_Parame = g_str_Parame & "               '001'                                                                                               AS CAMPO_DET_03, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_04, "
      g_str_Parame = g_str_Parame & "               'INTERES ' ||"
      g_str_Parame = g_str_Parame & "               SUBSTR(TO_CHAR(TO_DATE(PPGCAB_FECPPG,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(PPGCAB_FECPPG,1,4)  AS CAMPO_DET_05, "
      g_str_Parame = g_str_Parame & "               1.000                                                                                               AS CAMPO_DET_06, "
      g_str_Parame = g_str_Parame & "               'NIU'                                                                                               AS CAMPO_DET_07, "
      g_str_Parame = g_str_Parame & "               (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)                                                          AS CAMPO_DET_08, "
      g_str_Parame = g_str_Parame & "               1.000 * (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)                                                  AS CAMPO_DET_09, "
      g_str_Parame = g_str_Parame & "               '9998'                                                                                              AS CAMPO_DET_10_1, " '1000
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_DET_10_2, "
      g_str_Parame = g_str_Parame & "               '30'                                                                                                AS CAMPO_DET_10_3, "
      g_str_Parame = g_str_Parame & "               1.000 * (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)                                                  AS CAMPO_DET_10_4, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_10_5, "
      
      g_str_Parame = g_str_Parame & "               (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)                                                          AS CAMPO_DET_11, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_12, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_13, "
      g_str_Parame = g_str_Parame & "               '84121901'                                                                                          AS CAMPO_DET_14, "
      g_str_Parame = g_str_Parame & "               '7001'                                                                                              AS CAMPO_DET_15_1_1, "
      g_str_Parame = g_str_Parame & "               '1'                                                                                                 AS CAMPO_DET_15_2_1, "
      g_str_Parame = g_str_Parame & "               '7002'                                                                                              AS CAMPO_DET_15_1_2, "
      g_str_Parame = g_str_Parame & "               CASE WHEN L.SOLMAE_PRIVIV = 1 THEN 3  "
      g_str_Parame = g_str_Parame & "                    WHEN L.SOLMAE_PRIVIV = 2 THEN 0 END                                                            AS CAMPO_DET_15_2_2, "
      g_str_Parame = g_str_Parame & "               '7003'                                                                                              AS CAMPO_DET_15_1_3, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_15_2_3, "
      g_str_Parame = g_str_Parame & "               '7004'                                                                                              AS CAMPO_DET_15_1_4, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_15_2_4, "
      g_str_Parame = g_str_Parame & "               '7005'                                                                                              AS CAMPO_DET_15_1_5, "
      g_str_Parame = g_str_Parame & "               SUBSTR(J.HIPMAE_FECACT,1,4) || '-' || SUBSTR(J.HIPMAE_FECACT,5,2) || '-' || SUBSTR(J.HIPMAE_FECACT,7,2) "
      g_str_Parame = g_str_Parame & "                                                                                                                   AS CAMPO_DET_15_2_5, "
      g_str_Parame = g_str_Parame & "               '7006'                                                                                              AS CAMPO_DET_15_1_6, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_15_2_6, "
      g_str_Parame = g_str_Parame & "               '7007'                                                                                              AS CAMPO_DET_15_1_7, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_15_2_7, "
      g_str_Parame = g_str_Parame & "               '7008'                                                                                              AS CAMPO_DET_15_1_8, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_15_2_8, "
      g_str_Parame = g_str_Parame & "               '7009'                                                                                              AS CAMPO_DET_15_1_9, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_15_2_9, "
      g_str_Parame = g_str_Parame & "               '7010'                                                                                              AS CAMPO_DET_15_1_10, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_15_2_10, "
      g_str_Parame = g_str_Parame & "               '7011'                                                                                              AS CAMPO_DET_15_1_11, "
      g_str_Parame = g_str_Parame & "               '-'                                                                                                 AS CAMPO_DET_15_2_11, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_37, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_DET_38, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_39_1, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_39_2, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_39_3, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_39_4, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET_39_5, "
      
      g_str_Parame = g_str_Parame & "               'DET2'                                                                                              AS CAMPO_DET2_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_DET2_02, "
      g_str_Parame = g_str_Parame & "               '002'                                                                                               AS CAMPO_DET2_03, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_04, "
      g_str_Parame = g_str_Parame & "               'OTROS IMPORTES ' || "
      g_str_Parame = g_str_Parame & "               SUBSTR(TO_CHAR(TO_DATE(A.PPGCAB_FECPPG,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(PPGCAB_FECPPG,1,4)  AS CAMPO_DET2_05, "
      g_str_Parame = g_str_Parame & "               1.000                                                                                               AS CAMPO_DET2_06, "
      g_str_Parame = g_str_Parame & "               'NIU'                                                                                               AS CAMPO_DET2_07, "
      g_str_Parame = g_str_Parame & "               ((CASE WHEN A.PPGCAB_TIPPPG = 1 THEN A.PPGCAB_MTODEP ELSE PPGCAB_MTOTOT END) - (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)) "
      g_str_Parame = g_str_Parame & "                                                                                                                   AS CAMPO_DET2_08, "
      g_str_Parame = g_str_Parame & "               1.000 * ((CASE WHEN A.PPGCAB_TIPPPG = 1 THEN A.PPGCAB_MTODEP ELSE PPGCAB_MTOTOT END) - (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)) "
      g_str_Parame = g_str_Parame & "                                                                                                                   AS CAMPO_DET2_09, "
      g_str_Parame = g_str_Parame & "               '9998'                                                                                              AS CAMPO_DET2_10_1, " '1000
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_DET2_10_2, "
      g_str_Parame = g_str_Parame & "               '30'                                                                                                AS CAMPO_DET2_10_3, "
      g_str_Parame = g_str_Parame & "               1.000 * (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC)                                                  AS CAMPO_DET2_10_4, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_10_5, "
      
      g_str_Parame = g_str_Parame & "               (CASE WHEN A.PPGCAB_TIPPPG = 1 THEN A.PPGCAB_MTODEP ELSE PPGCAB_MTOTOT END) - (A.PPGCAB_INTCAL_TNC + A.PPGCAB_INTCAL_TC) "
      g_str_Parame = g_str_Parame & "                                                                                                                   AS CAMPO_DET2_11, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_12, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_13, "
      g_str_Parame = g_str_Parame & "               '84121501'                                                                                          AS CAMPO_DET2_14, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_1, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_1, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_2, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_2, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_3, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_3, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_4, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_4, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_5, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_5, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_6, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_6, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_7, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_7, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_8, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_8, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_9, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_9, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_10, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_10, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_1_11, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_15_2_11, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_37, "
      g_str_Parame = g_str_Parame & "               0.00                                                                                                AS CAMPO_DET2_38,"
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_39_1,"
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_39_2,"
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_39_3,"
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_39_4,"
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_DET2_39_5,"

      g_str_Parame = g_str_Parame & "               'ADI1'                                                                                              AS CAMPO_ADI_01, "
      g_str_Parame = g_str_Parame & "               'B'                                                                                                 AS CAMPO_ADI_02, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_ADI_03, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS CAMPO_ADI_04, "
      g_str_Parame = g_str_Parame & "               TRIM(A.PPGCAB_NUMOPE)                                                                               AS OPERACION, "
      g_str_Parame = g_str_Parame & "               ''                                                                                                  AS NUMERO_MOVIMIENTO, "
      g_str_Parame = g_str_Parame & "               HIPMAE_FECCAN                                                                                       AS FECHA_CANCELACION, "
      g_str_Parame = g_str_Parame & "               HIPMAE_SITUAC                                                                                       AS SITUACION, "
      g_str_Parame = g_str_Parame & "               PPGCAB_FECPPG                                                                                       AS FECHA_DEPOSITO "
      g_str_Parame = g_str_Parame & "          FROM CRE_PPGCAB A "
      g_str_Parame = g_str_Parame & "               INNER JOIN CRE_HIPMAE J ON J.HIPMAE_NUMOPE = A.PPGCAB_NUMOPE "
      g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '204' AND B.PARDES_CODITE = J.HIPMAE_MONEDA "
      g_str_Parame = g_str_Parame & "               INNER JOIN CNTBL_CATSUN C ON C.CATSUN_NROCAT = 2 AND C.CATSUN_DESCRI = TRIM(B.PARDES_DESCRI) "
      g_str_Parame = g_str_Parame & "               INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = J.HIPMAE_TDOCLI AND D.DATGEN_NUMDOC = J.HIPMAE_NDOCLI"
      g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 201 AND E.PARDES_CODITE = D.DATGEN_TIPVIA "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 202 AND F.PARDES_CODITE = D.DATGEN_TIPZON "
      g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES G ON G.PARDES_CODGRP = 101 AND G.PARDES_CODITE = D.DATGEN_UBIGEO "
      g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 101 AND H.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,2)||'0000' "
      g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES I ON I.PARDES_CODGRP = 101 AND I.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "                LEFT JOIN CRE_SOLMAE L ON L.SOLMAE_NUMERO = J.HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & "                LEFT JOIN CRE_SOLINM P ON P.SOLINM_NUMSOL = J.HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES Q ON Q.PARDES_CODGRP = 201 AND Q.PARDES_CODITE = P.SOLINM_TIPVIA "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES R ON R.PARDES_CODGRP = 202 AND R.PARDES_CODITE = P.SOLINM_TIPZON "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES S ON S.PARDES_CODGRP = 101 AND S.PARDES_CODITE = P.SOLINM_UBIGEO "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES T ON T.PARDES_CODGRP = 101 AND T.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,4)||'00' "
      g_str_Parame = g_str_Parame & "                LEFT JOIN MNT_PARDES U ON U.PARDES_CODGRP = 101 AND U.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,2)||'0000' "
      g_str_Parame = g_str_Parame & "         WHERE "
      g_str_Parame = g_str_Parame & "               A.PPGCAB_FECPPG >= 20180801 "
      g_str_Parame = g_str_Parame & "           AND A.PPGCAB_FECPPG <= 20190331 "
      g_str_Parame = g_str_Parame & "           AND A.PPGCAB_FLGPRO = 0 "
      g_str_Parame = g_str_Parame & "           AND HIPMAE_TDOCLI IN (1, 6) "
      'g_str_Parame = g_str_Parame & "           AND HIPMAE_NUMOPE IN ('0071300013') "
      g_str_Parame = g_str_Parame & "         ORDER BY A.PPGCAB_FECPPG "

   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta principal, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontró ningún registro.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro anterior en OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
    
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
      moddat_g_str_NumOpe = g_rst_Princi!OPERACION
      moddat_g_str_Codigo = g_rst_Princi!NUMERO_MOVIMIENTO
      
      If g_rst_Princi!SITUACION <> 2 Then
         If g_rst_Princi!FECHA_DEPOSITO <= g_rst_Princi!FECHA_CANCELACION Then
            GoTo Ingresar
         End If
      Else
      
Ingresar:
         Call fs_Obtener_Codigo("03", r_lng_Contad, r_int_SerFac, r_lng_NumFac)
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELE (      "
         g_str_Parame = g_str_Parame & " DOCELE_CODIGO                 , "
         g_str_Parame = g_str_Parame & " DOCELE_NUMOPE                 , "
         g_str_Parame = g_str_Parame & " DOCELE_NUMMOV                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FECPRO                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FECAUT                 , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_FECEMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_HOREMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPMON             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_NUMORC             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_FECVCT             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_NOMCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_UBIGEO             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_URBANI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_TELEMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_COREMI             , "
         
         g_str_Parame = g_str_Parame & " DOCELE_EMI_CODSUN             , "
         
         g_str_Parame = g_str_Parame & " DOCELE_REC_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_TELREC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_CORREC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_CODMOT             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_DESMOT             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXP      , "
         
'         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_INDICA_CARDSC      , "
         
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_REGPER      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_BASIMP_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOPER             , "
         
'         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOTOT_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_FACTOR_TASPER      , "
         
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIMP             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOIMP             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_OTRCAR             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_TOTDSC      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTDSC             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_IMPTOT_DOCUME      , "

'         g_str_Parame = g_str_Parame & " DOCELE_CAB_DSCGLO             , "

         g_str_Parame = g_str_Parame & " DOCELE_CAB_INFPPG             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTANT             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TIPOPE             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_LEYEND             , "
         
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOTOT_IMP         , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CARDSC             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODMOT_CARDSC      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_FACTOR_CARDSC      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MONTO_CARDSC       , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOBAS_CARDSC      , "
         
         g_str_Parame = g_str_Parame & " DOCELE_ADI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_TITADI             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_VALADI             , "
         g_str_Parame = g_str_Parame & " DOCELE_FLGENV                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FLGRPT                 , "
         g_str_Parame = g_str_Parame & " DOCELE_SITUAC                 , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES ( "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & " , "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
         g_str_Parame = g_str_Parame & "" & moddat_g_str_Codigo & " , "
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "     , "
         g_str_Parame = g_str_Parame & " NULL, "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_05 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_06 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_05 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_06 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_09 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_10 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_11 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_12 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_13 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_14 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_15 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_16 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_05 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_REC_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "'" & Mid(Replace(g_rst_Princi!CAMPO_REC_06, "  ", " "), 1, 100) & "'                                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_09 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_10 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_11 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_12 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_04 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DRF_05) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DRF_05 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_06 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_03 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_04) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_04 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_06 & "                    , "
         End If
               
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_08 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_09 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_10) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_10 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_11 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_12 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_13 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_14 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_15) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_15 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_16) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_16 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_17) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_17 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_18_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_18_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_18_2 & "                  , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_19) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_19 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_20 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_21 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_22) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_22 & "                    , "
         End If
         
'         If IsNull(g_rst_Princi!CAMPO_CAB_23) Then
'            g_str_Parame = g_str_Parame & " NULL, "
'         Else
'            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_23 & "                    , "
'         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_23 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_24) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_24 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_25 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_26 & "'                     , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_27 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_28_1 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_28_2 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_28_3 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_28_4 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_28_5 & "'                   , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_04 & "'                     , "
         g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
         g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
         g_str_Parame = g_str_Parame & "" & 1 & "                                               , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                               
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELE, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
   
         
         ''INTERES COMPENSATORIO
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOBAS_IMP      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TASTRI          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         
'         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "

         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD_GS1      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOTOT_IMP      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_INDCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FACCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOBAS_CAR      , "
         
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES (                        "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_04 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_06 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_08 & "                    , "
         End If
   
         If IsNull(g_rst_Princi!CAMPO_DET_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_09 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_10_2 & "                  , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_3 & "'                   , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_4 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_5 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_11 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_12 & "                    , "
         End If
      
'         If IsNull(g_rst_Princi!CAMPO_DET_13) Then
'            g_str_Parame = g_str_Parame & " NULL, "
'         Else
'            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_13 & "                    , "
'         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_13 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_14 & "'                     , "

         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_1 & "'                 , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_15_2_1) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_15_2_1 & "                , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_2 & "'                 , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_15_2_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_15_2_2 & "                , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_3 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_3 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_4 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_4 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_5 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_5 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_6 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_6 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_7 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_7 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_8 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_8 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_9 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_9 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_10 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_10 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_1_11 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15_2_11 & "'                , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_37 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_38 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_39_1 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_39_2 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_39_3 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_39_4 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_39_5 & "'                   , "
         
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "'                           , "
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de INTERES COMPENSATORIO, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
         
                                      
         ''OTROS IMPORTES
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOBAS_IMP      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TASTRI          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         
'         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "

         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD_GS1      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOTOT_IMP      , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_INDCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FACCAR          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOBAS_CAR      , "
         
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES (                        "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_04 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_06 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_08 & "                    , "
         End If
   
         If IsNull(g_rst_Princi!CAMPO_DET2_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_09 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_10_2 & "                  , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_3 & "'                   , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_4 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_5 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_11 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_12 & "                    , "
         End If
      
'         If IsNull(g_rst_Princi!CAMPO_DET2_13) Then
'            g_str_Parame = g_str_Parame & " NULL, "
'         Else
'            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_13 & "                    , "
'         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_13 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_14 & "'                     , "

         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_1 & "'                 , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_15_2_1) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_15_2_1 & "                , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_2 & "'                 , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_15_2_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_15_2_2 & "                , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_3 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_3 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_4 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_4 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_5 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_5 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_6 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_6 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_7 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_7 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_8 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_8 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_9 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_9 & "'                 , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_10 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_10 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_1_11 & "'                , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15_2_11 & "'                , "
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_37 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_38 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_39_1 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_39_2 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_39_3 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_39_4 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_39_5 & "'                   , "
         
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de OTROS IMPORTES, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
                                                                                                   
         DoEvents: DoEvents: DoEvents
            
         'ACTUALIZA EL CAMPO CAJMOV_FLGPRO PARA IDENTIFICAR CUALES SE HAN PROCESADO Y YA SE ENCUENTRAN EN CNTBL_DOCELE
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "UPDATE OPE_CAJMOV SET CAJMOV_FLGPRO = 1 "
         g_str_Parame = g_str_Parame & " WHERE CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' "
         g_str_Parame = g_str_Parame & "   AND CAJMOV_NUMMOV = '" & moddat_g_str_Codigo & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar CAJMOV_FLGPRO de la tabla OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
      
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Genera_FactAnterior")

End Sub
Private Sub fs_Generar_Facturas_ORI()
Dim r_lng_Contad     As Long
Dim r_int_SerFac     As Integer
Dim r_lng_NumFac     As Long

   On Error GoTo MyError
   
   Screen.MousePointer = 11
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT  CAMPO_IDE_01, CAMPO_IDE_02  , CAMPO_IDE_03 , CAMPO_IDE_04 , CAMPO_IDE_05   , CAMPO_IDE_06   , CAMPO_IDE_07   , CAMPO_IDE_08 , CAMPO_EMI_01  , CAMPO_EMI_02,"
   g_str_Parame = g_str_Parame & "           CAMPO_EMI_03, CAMPO_EMI_04  , CAMPO_EMI_05 , CAMPO_EMI_06 , CAMPO_EMI_07   , CAMPO_EMI_08   , CAMPO_EMI_09   , CAMPO_EMI_10 , CAMPO_EMI_11  , CAMPO_EMI_12,"
   g_str_Parame = g_str_Parame & "           CAMPO_EMI_13, CAMPO_EMI_14  , CAMPO_EMI_15 , CAMPO_REC_01 , CAMPO_REC_02   , CAMPO_REC_03   , CAMPO_REC_04   , CAMPO_REC_05 , CAMPO_REC_06  , CAMPO_REC_07,"
   g_str_Parame = g_str_Parame & "           CAMPO_REC_08, CAMPO_REC_09  , CAMPO_REC_10 , CAMPO_REC_11 , CAMPO_REC_12   , CAMPO_DRF_01   , CAMPO_DRF_02   , CAMPO_DRF_03 , CAMPO_DRF_04  , CAMPO_DRF_05,"
   g_str_Parame = g_str_Parame & "           CAMPO_DRF_06, CAMPO_CAB_01  , CAMPO_CAB_02 , CAMPO_CAB_03 , CAMPO_CAB_04   , CAMPO_CAB_05   , CAMPO_CAB_06   , CAMPO_CAB_07 , CAMPO_CAB_08  , CAMPO_CAB_09,"
   g_str_Parame = g_str_Parame & "           CAMPO_CAB_10, CAMPO_CAB_11  , CAMPO_CAB_12 , CAMPO_CAB_13 , CAMPO_CAB_14   , CAMPO_CAB_15   , CAMPO_CAB_16   , CAMPO_CAB_17 , CAMPO_CAB_18_1, CAMPO_CAB_18_2,"
   g_str_Parame = g_str_Parame & "           CAMPO_CAB_19, CAMPO_CAB_20  , CAMPO_CAB_21 , CAMPO_CAB_22 , CAMPO_CAB_23   , CAMPO_CAB_24   , CAMPO_CAB_25   , CAMPO_CAB_26 , CAMPO_CAB_27  , CAMPO_DET_01,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_02, CAMPO_DET_03  , CAMPO_DET_04 , "
   g_str_Parame = g_str_Parame & "           MIN(CAMPO_DET_05) AS CAMPO_DET_05, "
   g_str_Parame = g_str_Parame & "           CAMPO_DET_06, CAMPO_DET_07  , CAMPO_DET_08 , CAMPO_DET_09 , CAMPO_DET_10_1 , CAMPO_DET_10_2 , CAMPO_DET_10_3 , CAMPO_DET_11 , CAMPO_DET_12  , CAMPO_DET_13,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_14, CAMPO_DET_15  , CAMPO_DET_16 , CAMPO_DET_17 , CAMPO_DET_18   , CAMPO_DET_19   , CAMPO_DET_20   , CAMPO_DET_21 , CAMPO_DET_22  , CAMPO_DET_23,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_24, CAMPO_DET_25  , CAMPO_DET_26 , CAMPO_DET_27 , CAMPO_DET_28   , CAMPO_DET_29   , CAMPO_DET_30   , CAMPO_DET_31 , CAMPO_DET_32  , CAMPO_DET_33,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET_34, CAMPO_DET_35  , CAMPO_DET_36 , CAMPO_DET_37 , CAMPO_DET2_01  , CAMPO_DET2_02  , CAMPO_DET2_03  , CAMPO_DET2_04,"
   g_str_Parame = g_str_Parame & "           MIN(CAMPO_DET2_05) AS CAMPO_DET2_05, "
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_06, CAMPO_DET2_07, CAMPO_DET2_08, CAMPO_DET2_09, CAMPO_DET2_10_1, CAMPO_DET2_10_2, CAMPO_DET2_10_3, CAMPO_DET2_11, CAMPO_DET2_12 , CAMPO_DET2_13,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_14, CAMPO_DET2_15, CAMPO_DET2_16, CAMPO_DET2_17, CAMPO_DET2_18  , CAMPO_DET2_19  , CAMPO_DET2_20  , CAMPO_DET2_21, CAMPO_DET2_22 , CAMPO_DET2_23,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_24, CAMPO_DET2_25, CAMPO_DET2_26, CAMPO_DET2_27, CAMPO_DET2_28  , CAMPO_DET2_29  , CAMPO_DET2_30  , CAMPO_DET2_31, CAMPO_DET2_32 , CAMPO_DET2_33,"
   g_str_Parame = g_str_Parame & "           CAMPO_DET2_34, CAMPO_DET2_35, CAMPO_DET2_36, CAMPO_DET2_37, CAMPO_ADI_01   , CAMPO_ADI_02   , CAMPO_ADI_03   , CAMPO_ADI_04 , OPERACION     , NUMERO_MOVIMIENTO,"
   g_str_Parame = g_str_Parame & "           FECHA_CANCELACION , SITUACION, FECHA_DEPOSITO "
   g_str_Parame = g_str_Parame & "     FROM ( "
'''
   g_str_Parame = g_str_Parame & "     SELECT 'IDE'                                                                                               AS CAMPO_IDE_01, "
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_IDE_02, "  '001-
   g_str_Parame = g_str_Parame & "            SUBSTR(CAJMOV_FECDEP,1,4) || '-' || SUBSTR(CAJMOV_FECDEP,5,2) || '-' || SUBSTR(CAJMOV_FECDEP,7,2)   AS CAMPO_IDE_03, "
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_04,"
   g_str_Parame = g_str_Parame & "            '01'                                                                                                AS CAMPO_IDE_05,              " '--CATALOGO N°1.- FACTURA
   g_str_Parame = g_str_Parame & "            C.CATSUN_CODIGO                                                                                     AS CAMPO_IDE_06,              " '--CATALOGO N°2.-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_07,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_08,"
   g_str_Parame = g_str_Parame & "            'EMI'                                                                                               AS CAMPO_EMI_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_EMI_02,"  '001-
   g_str_Parame = g_str_Parame & "            '6'                                                                                                 AS CAMPO_EMI_03,              " '--CATALOGO N°6.-
   g_str_Parame = g_str_Parame & "            '20511904162'                                                                                       AS CAMPO_EMI_04,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_05,"
   g_str_Parame = g_str_Parame & "            'EDPYME MICASITA SA'                                                                                AS CAMPO_EMI_06,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_07,              " '--ATALOGO N°13.-UBIGEO
   g_str_Parame = g_str_Parame & "            'AV RIVERA NAVARRETE 645'                                                                           AS CAMPO_EMI_08,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_09,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_10,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_11,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_12,"
   g_str_Parame = g_str_Parame & "            'PE'                                                                                                AS CAMPO_EMI_13,              " '--CATALOGO N°4.-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_14,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_15,"
   g_str_Parame = g_str_Parame & "            'REC'                                                                                               AS CAMPO_REC_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_REC_02,"  '001-
   g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_TIPDOC)                                                                               AS CAMPO_REC_03,              " '--CATALOGO N°6.-
   g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_NUMDOC)                                                                               AS CAMPO_REC_04,"
   g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_APEPAT) || ' ' || TRIM(D.DATGEN_APEMAT) || ' ' || TRIM(D.DATGEN_NOMBRE)               AS CAMPO_REC_05,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_REC_06,"
   g_str_Parame = g_str_Parame & "            TRIM(H.PARDES_DESCRI)                                                                               AS CAMPO_REC_07,"
   g_str_Parame = g_str_Parame & "            TRIM(I.PARDES_DESCRI)                                                                               AS CAMPO_REC_08,"
   g_str_Parame = g_str_Parame & "            TRIM(G.PARDES_DESCRI)                                                                               AS CAMPO_REC_09,"
   g_str_Parame = g_str_Parame & "            'PE'                                                                                                AS CAMPO_REC_10,"
   g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_TELEFO)                                                                               AS CAMPO_REC_11,"
   g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_DIRELE)                                                                               AS CAMPO_REC_12,"
   g_str_Parame = g_str_Parame & "            'DRF'                                                                                               AS CAMPO_DRF_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_DRF_02," '001-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_03,               "
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_04,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_05,               " '--PARA NOTA CREDITO/DEBITO
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_06,"
   g_str_Parame = g_str_Parame & "            'CAB'                                                                                               AS CAMPO_CAB_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_CAB_02," '001-
   g_str_Parame = g_str_Parame & "            '1001'                                                                                              AS CAMPO_CAB_03,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_04,"
   g_str_Parame = g_str_Parame & "            '1002'                                                                                              AS CAMPO_CAB_05,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_06,"
   g_str_Parame = g_str_Parame & "            '1003'                                                                                              AS CAMPO_CAB_07,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            '0.00'                                                                                              AS CAMPO_CAB_08,"
   g_str_Parame = g_str_Parame & "            '1004'                                                                                              AS CAMPO_CAB_09,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            '0.00'                                                                                              AS CAMPO_CAB_10,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_CAB_11,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_12,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_13,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_14,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_15,"
   
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_16,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_17,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_CAB_18_1,"
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_18_2,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_19,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_20,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_21,"
   g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_22,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_23,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_24,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_25,"
   g_str_Parame = g_str_Parame & "            '13'                                                                                                AS CAMPO_CAB_26,                " '--CATALOGO N°17.-
   g_str_Parame = g_str_Parame & "            '[1000'                                                                                             AS CAMPO_CAB_27,                " '--CATALOGO N°15.- DETALLE EN LETRAS DEL IMPORTE
   
   g_str_Parame = g_str_Parame & "            'DET1'                                                                                              AS CAMPO_DET_01,"
'   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_DET_02," '001-
   g_str_Parame = g_str_Parame & "            '001'                                                                                               AS CAMPO_DET_03,                " '-- Número de orden de ítem
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_04,                "
   g_str_Parame = g_str_Parame & "            'INTERES ' ||"
   g_str_Parame = g_str_Parame & "            SUBSTR(TO_CHAR(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(HIPCUO_FECVCT,1,4)  AS CAMPO_DET_05," 'CAJMOV_FECDEP
   g_str_Parame = g_str_Parame & "            1.000                                                                                               AS CAMPO_DET_06,"
   g_str_Parame = g_str_Parame & "            'NIU'                                                                                               AS CAMPO_DET_07,                " '--CATALOGO N°3.-
   g_str_Parame = g_str_Parame & "            A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_08,"
   g_str_Parame = g_str_Parame & "            1.000 * A.CAJMOV_INTERE                                                                             AS CAMPO_DET_09,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_DET_10_1,              " '--CATALOGO N°5,7 u 8.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_DET_10_2,"
   g_str_Parame = g_str_Parame & "            '30'                                                                                                AS CAMPO_DET_10_3,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_10_4,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_10_5,"
   
   g_str_Parame = g_str_Parame & "            A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_11,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_12,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_13,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_14,"
   g_str_Parame = g_str_Parame & "            '84121901'                                                                                          AS CAMPO_DET_15,                " '--CATALOGO N°15.-
   
   
   g_str_Parame = g_str_Parame & "            '7004'                                                                                              AS CAMPO_DET_16,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_17,                " '--NRO CONTRATO O PRESTAMO
   g_str_Parame = g_str_Parame & "            '7005'                                                                                              AS CAMPO_DET_18,"
   g_str_Parame = g_str_Parame & "            SUBSTR(J.HIPMAE_FECACT,1,4) || '-' || SUBSTR(J.HIPMAE_FECACT,5,2) || '-' || SUBSTR(J.HIPMAE_FECACT,7,2)"
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET_19,"
   g_str_Parame = g_str_Parame & "            '7001'                                                                                              AS CAMPO_DET_20,"
   g_str_Parame = g_str_Parame & "            '1'                                                                                                 AS CAMPO_DET_21,                " '--CATALOGO N°26.- Si es construcción /adquisición
   g_str_Parame = g_str_Parame & "            '7003'                                                                                              AS CAMPO_DET_22,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_23,"
   g_str_Parame = g_str_Parame & "            '7002'                                                                                              AS CAMPO_DET_24,"
   g_str_Parame = g_str_Parame & "            CASE WHEN L.SOLMAE_PRIVIV = 1 THEN 3  "
   g_str_Parame = g_str_Parame & "                 WHEN L.SOLMAE_PRIVIV = 2 THEN 0 END                                                            AS CAMPO_DET_25,                " '--CATALOGO N°27.- VERIFICAR
   g_str_Parame = g_str_Parame & "            '7007'                                                                                              AS CAMPO_DET_26,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_27,"
   g_str_Parame = g_str_Parame & "            '7006'                                                                                              AS CAMPO_DET_28,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_29,"                 '/*P.SOLINM_UBIGEO*/
   g_str_Parame = g_str_Parame & "            '7008'                                                                                              AS CAMPO_DET_30,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_31,"
   g_str_Parame = g_str_Parame & "            '7011'                                                                                              AS CAMPO_DET_32,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_33,"
   g_str_Parame = g_str_Parame & "            '7009'                                                                                              AS CAMPO_DET_34,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_35,"
   g_str_Parame = g_str_Parame & "            '7010'                                                                                              AS CAMPO_DET_36,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_37,"
   
   g_str_Parame = g_str_Parame & "            'DET2'                                                                                              AS CAMPO_DET2_01,"
'   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_DET2_02," '001-
   g_str_Parame = g_str_Parame & "            '002'                                                                                               AS CAMPO_DET2_03,                " '-- Número de orden de ítem
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_04,                "
   g_str_Parame = g_str_Parame & "            'OTROS IMPORTES ' ||"
   g_str_Parame = g_str_Parame & "            SUBSTR(TO_CHAR(TO_DATE(HIPCUO_FECVCT,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(HIPCUO_FECVCT,1,4)  AS CAMPO_DET2_05," 'CAJMOV_FECDEP
   g_str_Parame = g_str_Parame & "            1.000                                                                                               AS CAMPO_DET2_06,"
   g_str_Parame = g_str_Parame & "            'NIU'                                                                                               AS CAMPO_DET2_07,                " '--CATALOGO N°3.-
   g_str_Parame = g_str_Parame & "            (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)"
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_08,"
   g_str_Parame = g_str_Parame & "            1.000 * (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)"
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_09,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_DET2_10_1,              " '--CATALOGO N°5,7 u 8.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_DET2_10_2,"
   g_str_Parame = g_str_Parame & "            '30'                                                                                                AS CAMPO_DET2_10_3,"
   g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG - A.CAJMOV_INTERE "
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_11,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_12,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_13,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_14,"
   g_str_Parame = g_str_Parame & "            '84121501'                                                                                          AS CAMPO_DET2_15,                " '--CATALOGO N°15.-
   
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_16,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_17,                " '--NRO CONTRATO O PRESTAMO
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_18,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_19,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_20,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_21,                " '--CATALOGO N°26.- Si es construcción /adquisición
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_22,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_23,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_24,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_25,                " '--CATALOGO N°27.- VERIFICAR
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_26,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_27,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_28,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_29,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_30,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_31,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_32,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_33,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_34,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_35,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_36,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_37,"

   g_str_Parame = g_str_Parame & "            'ADI1'                                                                                              AS CAMPO_ADI_01,"
'   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_ADI_02," '001-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_ADI_03,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_ADI_04,"
   g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_NUMOPE)                                                                               AS OPERACION, "
   g_str_Parame = g_str_Parame & "            CAJMOV_NUMMOV                                                                                       AS NUMERO_MOVIMIENTO , "
   g_str_Parame = g_str_Parame & "            HIPMAE_FECCAN                                                                                       AS FECHA_CANCELACION ,"
   g_str_Parame = g_str_Parame & "            HIPMAE_SITUAC                                                                                       AS SITUACION, "
   g_str_Parame = g_str_Parame & "            CAJMOV_FECDEP                                                                                       AS FECHA_DEPOSITO "
   g_str_Parame = g_str_Parame & "       FROM OPE_CAJMOV A "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '204' AND B.PARDES_CODITE = A.CAJMOV_MONPAG "
   g_str_Parame = g_str_Parame & "            INNER JOIN CNTBL_CATSUN C ON C.CATSUN_NROCAT = 2 AND C.CATSUN_DESCRI = TRIM(B.PARDES_DESCRI) "
   g_str_Parame = g_str_Parame & "            INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = A.CAJMOV_TIPDOC AND D.DATGEN_NUMDOC = A.CAJMOV_NUMDOC "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 201 AND E.PARDES_CODITE = D.DATGEN_TIPVIA "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 202 AND F.PARDES_CODITE = D.DATGEN_TIPZON "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES G ON G.PARDES_CODGRP = 101 AND G.PARDES_CODITE = D.DATGEN_UBIGEO "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 101 AND H.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,2)||'0000' "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES I ON I.PARDES_CODGRP = 101 AND I.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,4)||'00' "
   g_str_Parame = g_str_Parame & "            INNER JOIN CRE_HIPMAE J ON J.HIPMAE_NUMOPE = A.CAJMOV_NUMOPE "
   g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_HIPGAR K ON K.HIPGAR_NUMOPE = J.HIPMAE_NUMOPE AND K.HIPGAR_BIEGAR = 1 "
   g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_SOLMAE L ON L.SOLMAE_NUMERO = J.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_SOLINM P ON P.SOLINM_NUMSOL = J.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES Q ON Q.PARDES_CODGRP = 201 AND Q.PARDES_CODITE = P.SOLINM_TIPVIA "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES R ON R.PARDES_CODGRP = 202 AND R.PARDES_CODITE = P.SOLINM_TIPZON "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES S ON S.PARDES_CODGRP = 101 AND S.PARDES_CODITE = P.SOLINM_UBIGEO "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES T ON T.PARDES_CODGRP = 101 AND T.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,4)||'00' "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES U ON U.PARDES_CODGRP = 101 AND U.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,2)||'0000' "
   
   g_str_Parame = g_str_Parame & "            INNER JOIN CRE_HIPPAG V ON V.HIPPAG_NUMOPE = A.CAJMOV_NUMOPE AND V.HIPPAG_FECPAG = A.CAJMOV_FECDEP AND V.HIPPAG_NUMMOV = A.CAJMOV_NUMMOV"
   g_str_Parame = g_str_Parame & "            INNER JOIN CRE_HIPCUO W ON W.HIPCUO_NUMOPE = V.HIPPAG_NUMOPE AND W.HIPCUO_TIPCRO = 1 AND W.HIPCUO_NUMCUO = V.HIPPAG_NUMCUO"
               
   g_str_Parame = g_str_Parame & "      WHERE CAJMOV_SUCMOV IS NOT NULL "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_USUMOV IS NOT NULL "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_FECMOV > 0"
   g_str_Parame = g_str_Parame & "        AND CAJMOV_NUMMOV > 0 "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_CODBAN IS NOT NULL "
   
   If Chk_FecAct.Value = 0 Then
      g_str_Parame = g_str_Parame & "     AND CAJMOV_FECDEP = '" & l_str_FecCar & "' "
   Else
      g_str_Parame = g_str_Parame & "     AND CAJMOV_FECDEP <= '" & l_str_FecCar & "' "
   End If
   g_str_Parame = g_str_Parame & "        AND CAJMOV_FLGPRO = 0 "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_TIPMOV = '1102' "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_TIPDOC IN (1, 6) "                      '1-DNI y 6-RUC
   g_str_Parame = g_str_Parame & "      ORDER BY A.CAJMOV_FECMOV , A.CAJMOV_NUMMOV "
   g_str_Parame = g_str_Parame & "      )"
   
   g_str_Parame = g_str_Parame & "      GROUP BY CAMPO_IDE_01 , CAMPO_IDE_02  , CAMPO_IDE_03 , CAMPO_IDE_04 , CAMPO_IDE_05   , CAMPO_IDE_06   , CAMPO_IDE_07   , CAMPO_IDE_08  , CAMPO_EMI_01  , CAMPO_EMI_02,"
   g_str_Parame = g_str_Parame & "               CAMPO_EMI_03 , CAMPO_EMI_04  , CAMPO_EMI_05 , CAMPO_EMI_06 , CAMPO_EMI_07   , CAMPO_EMI_08   , CAMPO_EMI_09   , CAMPO_EMI_10  , CAMPO_EMI_11  , CAMPO_EMI_12,"
   g_str_Parame = g_str_Parame & "               CAMPO_EMI_13 , CAMPO_EMI_14  , CAMPO_EMI_15 , CAMPO_REC_01 , CAMPO_REC_02   , CAMPO_REC_03   , CAMPO_REC_04   , CAMPO_REC_05  , CAMPO_REC_06  , CAMPO_REC_07,"
   g_str_Parame = g_str_Parame & "               CAMPO_REC_08 , CAMPO_REC_09  , CAMPO_REC_10 , CAMPO_REC_11 , CAMPO_REC_12   , CAMPO_DRF_01   , CAMPO_DRF_02   , CAMPO_DRF_03  , CAMPO_DRF_04  , CAMPO_DRF_05,"
   g_str_Parame = g_str_Parame & "               CAMPO_DRF_06 , CAMPO_CAB_01  , CAMPO_CAB_02 , CAMPO_CAB_03 , CAMPO_CAB_04   , CAMPO_CAB_05   , CAMPO_CAB_06   , CAMPO_CAB_07  , CAMPO_CAB_08  , CAMPO_CAB_09,"
   g_str_Parame = g_str_Parame & "               CAMPO_CAB_10 , CAMPO_CAB_11  , CAMPO_CAB_12 , CAMPO_CAB_13 , CAMPO_CAB_14   , CAMPO_CAB_15   , CAMPO_CAB_16   , CAMPO_CAB_17  , CAMPO_CAB_18_1, CAMPO_CAB_18_2,"
   g_str_Parame = g_str_Parame & "               CAMPO_CAB_19 , CAMPO_CAB_20  , CAMPO_CAB_21 , CAMPO_CAB_22 , CAMPO_CAB_23   , CAMPO_CAB_24   , CAMPO_CAB_25   , CAMPO_CAB_26  , CAMPO_CAB_27  , CAMPO_DET_01,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET_02 , CAMPO_DET_03  , CAMPO_DET_04 , "
   g_str_Parame = g_str_Parame & "               CAMPO_DET_06 , CAMPO_DET_07  , CAMPO_DET_08 , CAMPO_DET_09 , CAMPO_DET_10_1 , CAMPO_DET_10_2 , CAMPO_DET_10_3 , CAMPO_DET_11  , CAMPO_DET_12  , CAMPO_DET_13,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET_14 , CAMPO_DET_15  , CAMPO_DET_16 , CAMPO_DET_17 , CAMPO_DET_18   , CAMPO_DET_19   , CAMPO_DET_20   , CAMPO_DET_21  , CAMPO_DET_22  , CAMPO_DET_23,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET_24 , CAMPO_DET_25  , CAMPO_DET_26 , CAMPO_DET_27 , CAMPO_DET_28   , CAMPO_DET_29   , CAMPO_DET_30   , CAMPO_DET_31  , CAMPO_DET_32  , CAMPO_DET_33,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET_34 , CAMPO_DET_35  , CAMPO_DET_36 , CAMPO_DET_37 , CAMPO_DET2_01  , CAMPO_DET2_02  , CAMPO_DET2_03  , CAMPO_DET2_04 , "
   g_str_Parame = g_str_Parame & "               CAMPO_DET2_06, CAMPO_DET2_07, CAMPO_DET2_08 , CAMPO_DET2_09, CAMPO_DET2_10_1, CAMPO_DET2_10_2, CAMPO_DET2_10_3, CAMPO_DET2_11 , CAMPO_DET2_12 , CAMPO_DET2_13,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET2_14, CAMPO_DET2_15, CAMPO_DET2_16 , CAMPO_DET2_17, CAMPO_DET2_18  , CAMPO_DET2_19  , CAMPO_DET2_20  , CAMPO_DET2_21 , CAMPO_DET2_22 , CAMPO_DET2_23,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET2_24, CAMPO_DET2_25, CAMPO_DET2_26 , CAMPO_DET2_27, CAMPO_DET2_28  , CAMPO_DET2_29  , CAMPO_DET2_30  , CAMPO_DET2_31 , CAMPO_DET2_32 , CAMPO_DET2_33,"
   g_str_Parame = g_str_Parame & "               CAMPO_DET2_34, CAMPO_DET2_35, CAMPO_DET2_36 , CAMPO_DET2_37, CAMPO_ADI_01   , CAMPO_ADI_02   , CAMPO_ADI_03   , CAMPO_ADI_04  , OPERACION     , NUMERO_MOVIMIENTO,"
   g_str_Parame = g_str_Parame & "               FECHA_CANCELACION , SITUACION, FECHA_DEPOSITO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta principal, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontró ningún registro.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro anterior en OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
    
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
      moddat_g_str_NumOpe = g_rst_Princi!OPERACION
      moddat_g_str_Codigo = g_rst_Princi!NUMERO_MOVIMIENTO
      
      If g_rst_Princi!SITUACION <> 2 Then
         If g_rst_Princi!FECHA_DEPOSITO <= g_rst_Princi!FECHA_CANCELACION Then
            GoTo Ingresar
         End If
      Else
      
Ingresar:
         Call fs_Obtener_Codigo("01", r_lng_Contad, r_int_SerFac, r_lng_NumFac)
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELE (      "
         g_str_Parame = g_str_Parame & " DOCELE_CODIGO                 , "
         g_str_Parame = g_str_Parame & " DOCELE_NUMOPE                 , "
         g_str_Parame = g_str_Parame & " DOCELE_NUMMOV                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FECPRO                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FECAUT                 , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_FECEMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_HOREMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPMON             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_NUMORC             , "
         g_str_Parame = g_str_Parame & " DOCELE_IDE_FECVCT             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_NOMCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_UBIGEO             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_URBANI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_TELEMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_EMI_COREMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DENOMI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DIRCOM             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DEPART             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_PROVIN             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_DISTRI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_CODPAI             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_TELREC             , "
         g_str_Parame = g_str_Parame & " DOCELE_REC_CORREC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_TIPDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_NUMDOC             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_CODMOT             , "
         g_str_Parame = g_str_Parame & " DOCELE_DRF_DESMOT             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRV      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEINA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXO      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRA      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_REGPER      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_BASIMP_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOPER             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOTOT_PERCEP      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIMP             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOIMP             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_OTRCAR             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_TOTDSC      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTDSC             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_IMPTOT_DOCUME      , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_DSCGLO             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_INFPPG             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTANT             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_TIPOPE             , "
         g_str_Parame = g_str_Parame & " DOCELE_CAB_LEYEND             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_SERNUM             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_TITADI             , "
         g_str_Parame = g_str_Parame & " DOCELE_ADI_VALADI             , "
         g_str_Parame = g_str_Parame & " DOCELE_FLGENV                 , "
         g_str_Parame = g_str_Parame & " DOCELE_FLGRPT                 , "
         g_str_Parame = g_str_Parame & " DOCELE_SITUAC                 , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES ( "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & " , "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
         g_str_Parame = g_str_Parame & "" & moddat_g_str_Codigo & " , "
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "     , "
         g_str_Parame = g_str_Parame & " NULL, "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_05 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_06 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_05 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_06 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_09 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_10 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_11 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_12 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_13 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_14 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_15 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_04 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_05 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_REC_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "'" & Mid(Replace(g_rst_Princi!CAMPO_REC_06, "  ", " "), 1, 100) & "'                                    , "
         End If
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_07 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_08 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_09 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_10 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_11 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_12 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_03 & "'                   , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_04 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DRF_05) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DRF_05 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_06 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_03 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_04) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_04 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_06 & "                    , "
         End If
               
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_08 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_09 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_10) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_10 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_11 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_12 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_13 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_14 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_15) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_15 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_16) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_16 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_17) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_17 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_18_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_18_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_18_2 & "                  , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_19) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_19 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_20 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_21 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_22) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_22 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_CAB_23) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_23 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_24 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_CAB_25) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_25 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_26 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_27 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_04 & "'                     , "
         g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
         g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
         g_str_Parame = g_str_Parame & "" & 1 & "                                               , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                               
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELE, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
   
         
         ''INTERES COMPENSATORIO
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES (                        "
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_03 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_04 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_05 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_06 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_07 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_08 & "                    , "
         End If
   
         If IsNull(g_rst_Princi!CAMPO_DET_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_09 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_1 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_10_2 & "                  , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_3 & "'                   , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_11 & "                    , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_12 & "                    , "
         End If
      
         If IsNull(g_rst_Princi!CAMPO_DET_13) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_13 & "                    , "
         End If
   
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_14 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_16 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_17 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_18 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_19 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_20 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_21 & "   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_22 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_23 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_24 & "'                     , "
         
         If IsNull(g_rst_Princi!CAMPO_DET_25) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_25 & "                    , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_26 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_27 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_28 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_29 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_30 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_31 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_32 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_33 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_34 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_35 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_36 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_37 & "'                     , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "'                           , "
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de INTERES COMPENSATORIO, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
         DoEvents: DoEvents: DoEvents
         
                                      
         ''OTROS IMPORTES
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
         g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
         g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
         g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
         g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
         g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
         g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
         g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
         g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
         g_str_Parame = g_str_Parame & " VALUES ("
         g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_03 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_04 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_05 & "'                    , "
         If IsNull(g_rst_Princi!CAMPO_DET2_06) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_06 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_07 & "'                    , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_08) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_08 & "                   , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_09) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_09 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_1 & "'                  , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_10_2) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_10_2 & "                 , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_3 & "'                  , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_11) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_11 & "                   , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_12) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_12 & "                   , "
         End If
         
         If IsNull(g_rst_Princi!CAMPO_DET2_13) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_13 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_14 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_16 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_17 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_18 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_19 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_20 & "'                    , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_21) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_21 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_22 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_23 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_24 & "'                    , "
         
         If IsNull(g_rst_Princi!CAMPO_DET2_25) Then
            g_str_Parame = g_str_Parame & " NULL, "
         Else
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_25 & "                   , "
         End If
         
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_26 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_27 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_28 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_29 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_30 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_31 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_32 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_33 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_34 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_35 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_36 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_37 & "'                    , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
         g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
         g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de OTROS IMPORTES, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
                                                                                                   
   
         DoEvents: DoEvents: DoEvents
         
         
         'ACTUALIZA EL CAMPO CAJMOV_FLGPRO PARA IDENTIFICAR CUALES SE HAN PROCESADO Y YA SE ENCUENTRAN EN CNTBL_DOCELE
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "UPDATE OPE_CAJMOV SET CAJMOV_FLGPRO = 1 "
         g_str_Parame = g_str_Parame & " WHERE CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' "
         g_str_Parame = g_str_Parame & "   AND CAJMOV_NUMMOV = '" & moddat_g_str_Codigo & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar CAJMOV_FLGPRO de la tabla OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
            Exit Sub
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
      
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Genera_FactAnterior")

End Sub
Private Sub fs_Generar_Facturas_old()
Dim r_lng_Contad     As Long
Dim r_int_SerFac     As Integer
Dim r_lng_NumFac     As Long

   On Error GoTo MyError
   
   Screen.MousePointer = 11
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "     SELECT 'IDE'                                                                                               AS CAMPO_IDE_01, "
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_IDE_02, "  '001-
   g_str_Parame = g_str_Parame & "            SUBSTR(CAJMOV_FECDEP,1,4) || '-' || SUBSTR(CAJMOV_FECDEP,5,2) || '-' || SUBSTR(CAJMOV_FECDEP,7,2)   AS CAMPO_IDE_03, "
'   g_str_Parame = g_str_Parame & "            CASE WHEN LENGTH(A.CAJMOV_HORMOV) = 6 THEN SUBSTR(A.CAJMOV_HORMOV,1,2) || ':' || SUBSTR(A.CAJMOV_HORMOV,3,2) || ':' || SUBSTR(A.CAJMOV_HORMOV,5,2) "
'   g_str_Parame = g_str_Parame & "            ELSE TRIM(TO_CHAR(SUBSTR(A.CAJMOV_HORMOV,1,1),'00')) || ':' || SUBSTR(A.CAJMOV_HORMOV,2,2) || ':' || SUBSTR(A.CAJMOV_HORMOV,4,2)  "
'   g_str_Parame = g_str_Parame & "             END                                                                                                "
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_04,"
   g_str_Parame = g_str_Parame & "            '01'                                                                                                AS CAMPO_IDE_05,              " '--CATALOGO N°1.- FACTURA
   g_str_Parame = g_str_Parame & "            C.CATSUN_CODIGO                                                                                     AS CAMPO_IDE_06,              " '--CATALOGO N°2.-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_07,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_IDE_08,"
   g_str_Parame = g_str_Parame & "            'EMI'                                                                                               AS CAMPO_EMI_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_EMI_02,"  '001-
   g_str_Parame = g_str_Parame & "            '6'                                                                                                 AS CAMPO_EMI_03,              " '--CATALOGO N°6.-
   g_str_Parame = g_str_Parame & "            '20511904162'                                                                                       AS CAMPO_EMI_04,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_05,"
   g_str_Parame = g_str_Parame & "            'EDPYME MICASITA SA'                                                                                AS CAMPO_EMI_06,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_07,              " '--ATALOGO N°13.-UBIGEO
   g_str_Parame = g_str_Parame & "            'AV RIVERA NAVARRETE 645'                                                                           AS CAMPO_EMI_08,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_09,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_10,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_11,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_12,"
   g_str_Parame = g_str_Parame & "            'PE'                                                                                                AS CAMPO_EMI_13,              " '--CATALOGO N°4.-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_14,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_EMI_15,"
   g_str_Parame = g_str_Parame & "            'REC'                                                                                               AS CAMPO_REC_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_REC_02,"  '001-
   g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_TIPDOC)                                                                               AS CAMPO_REC_03,              " '--CATALOGO N°6.-
   g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_NUMDOC)                                                                               AS CAMPO_REC_04,"
   g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_APEPAT) || ' ' || TRIM(D.DATGEN_APEMAT) || ' ' || TRIM(D.DATGEN_NOMBRE)               AS CAMPO_REC_05,"
   'g_str_Parame = g_str_Parame & "            TRIM(DECODE(D.DATGEN_TIPVIA, 12, '', TRIM(E.PARDES_DESCRI))||' '||TRIM(D.DATGEN_NOMVIA)||' '||TRIM(D.DATGEN_NUMERO)||' '||DECODE(NVL(LENGTH(TRIM(D.DATGEN_INTDPT)), 0), 0, '', '('||TRIM(D.DATGEN_INTDPT)||')')||' '||DECODE(NVL(LENGTH(TRIM(D.DATGEN_NOMZON)),0), 0, '', '-'||DECODE(D.DATGEN_TIPZON, 12, '', TRIM(F.PARDES_DESCRI))||' '||TRIM(D.DATGEN_NOMZON)))"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_REC_06,"
   g_str_Parame = g_str_Parame & "            TRIM(H.PARDES_DESCRI)                                                                               AS CAMPO_REC_07,"
   g_str_Parame = g_str_Parame & "            TRIM(I.PARDES_DESCRI)                                                                               AS CAMPO_REC_08,"
   g_str_Parame = g_str_Parame & "            TRIM(G.PARDES_DESCRI)                                                                               AS CAMPO_REC_09,"
   g_str_Parame = g_str_Parame & "            'PE'                                                                                                AS CAMPO_REC_10,"
   g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_TELEFO)                                                                               AS CAMPO_REC_11,"
   g_str_Parame = g_str_Parame & "            TRIM(D.DATGEN_DIRELE)                                                                               AS CAMPO_REC_12,"
   g_str_Parame = g_str_Parame & "            'DRF'                                                                                               AS CAMPO_DRF_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_DRF_02," '001-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_03,               "
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_04,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_05,               " '--PARA NOTA CREDITO/DEBITO
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DRF_06,"
   g_str_Parame = g_str_Parame & "            'CAB'                                                                                               AS CAMPO_CAB_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_CAB_02," '001-
   g_str_Parame = g_str_Parame & "            '1001'                                                                                              AS CAMPO_CAB_03,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_04,"
   'g_str_Parame = g_str_Parame & "            (A.CAJMOV_SEGDES + A.CAJMOV_SEGVIV) - (A.CAJMOV_SEGDES + A.CAJMOV_SEGVIV)*0.18                      AS CAMPO_CAB_04,"
   g_str_Parame = g_str_Parame & "            '1002'                                                                                              AS CAMPO_CAB_05,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_06,"
   'g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG - A.CAJMOV_SEGDES - A.CAJMOV_SEGVIV                                                 AS CAMPO_CAB_06,"
   g_str_Parame = g_str_Parame & "            '1003'                                                                                              AS CAMPO_CAB_07,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            '0.00'                                                                                              AS CAMPO_CAB_08,"
   g_str_Parame = g_str_Parame & "            '1004'                                                                                              AS CAMPO_CAB_09,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            '0.00'                                                                                              AS CAMPO_CAB_10,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_CAB_11,               " '--CATALOGO N°14.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_12,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_13,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_14,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_15,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_16,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_17,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_CAB_18_1,"
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_CAB_18_2,"
   'g_str_Parame = g_str_Parame & "            (A.CAJMOV_SEGDES + A.CAJMOV_SEGVIV) * 0.18                                                          AS CAMPO_CAB_18_2,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_19,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_20,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_21,"
   g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG                                                                                     AS CAMPO_CAB_22,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_23,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_24,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_CAB_25,"
   g_str_Parame = g_str_Parame & "            '13'                                                                                                AS CAMPO_CAB_26,                " '--CATALOGO N°17.-
   g_str_Parame = g_str_Parame & "            '[1000'                                                                                             AS CAMPO_CAB_27,                " '--CATALOGO N°15.- DETALLE EN LETRAS DEL IMPORTE
   g_str_Parame = g_str_Parame & "            'DET1'                                                                                              AS CAMPO_DET_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_DET_02," '001-
   g_str_Parame = g_str_Parame & "            '001'                                                                                               AS CAMPO_DET_03,                " '-- Número de orden de ítem
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_04,                "
   g_str_Parame = g_str_Parame & "            'INTERES ' ||"
   g_str_Parame = g_str_Parame & "            SUBSTR(TO_CHAR(TO_DATE(CAJMOV_FECMOV,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(CAJMOV_FECMOV,1,4)  AS CAMPO_DET_05,"
   g_str_Parame = g_str_Parame & "            1.000                                                                                               AS CAMPO_DET_06,"
   g_str_Parame = g_str_Parame & "            'NIU'                                                                                               AS CAMPO_DET_07,                " '--CATALOGO N°3.-
   g_str_Parame = g_str_Parame & "            A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_08,"
   g_str_Parame = g_str_Parame & "            1.000 * A.CAJMOV_INTERE                                                                             AS CAMPO_DET_09,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_DET_10_1,              " '--CATALOGO N°5,7 u 8.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_DET_10_2,"
   g_str_Parame = g_str_Parame & "            '30'                                                                                                AS CAMPO_DET_10_3,"
   g_str_Parame = g_str_Parame & "            A.CAJMOV_INTERE                                                                                     AS CAMPO_DET_11,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_12,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_13,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_14,"
   g_str_Parame = g_str_Parame & "            '84121901'                                                                                          AS CAMPO_DET_15,                " '--CATALOGO N°15.-
   g_str_Parame = g_str_Parame & "            '7004'                                                                                              AS CAMPO_DET_16,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET_17,                " '--NRO CONTRATO O PRESTAMO
   g_str_Parame = g_str_Parame & "            '7005'                                                                                              AS CAMPO_DET_18,"
   g_str_Parame = g_str_Parame & "            SUBSTR(J.HIPMAE_FECACT,1,4) || '-' || SUBSTR(J.HIPMAE_FECACT,5,2) || '-' || SUBSTR(J.HIPMAE_FECACT,7,2)"
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET_19,"
   g_str_Parame = g_str_Parame & "            '7001'                                                                                              AS CAMPO_DET_20,"
   g_str_Parame = g_str_Parame & "            '1'                                                                                                 AS CAMPO_DET_21,                " '--CATALOGO N°26.- Si es construcción /adquisición
   g_str_Parame = g_str_Parame & "            '7003'                                                                                              AS CAMPO_DET_22,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_23,"
   g_str_Parame = g_str_Parame & "            '7002'                                                                                              AS CAMPO_DET_24,"
   g_str_Parame = g_str_Parame & "            CASE WHEN L.SOLMAE_PRIVIV = 1 THEN 3  "
   g_str_Parame = g_str_Parame & "                 WHEN L.SOLMAE_PRIVIV = 2 THEN 0 END                                                            AS CAMPO_DET_25,                " '--CATALOGO N°27.- VERIFICAR
   g_str_Parame = g_str_Parame & "            '7007'                                                                                              AS CAMPO_DET_26,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_27,"
   g_str_Parame = g_str_Parame & "            '7006'                                                                                              AS CAMPO_DET_28,"
   g_str_Parame = g_str_Parame & "            '-'/*P.SOLINM_UBIGEO*/                                                                              AS CAMPO_DET_29,"
   g_str_Parame = g_str_Parame & "            '7008'                                                                                              AS CAMPO_DET_30,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_31,"
   g_str_Parame = g_str_Parame & "            '7011'                                                                                              AS CAMPO_DET_32,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_33,"
   g_str_Parame = g_str_Parame & "            '7009'                                                                                              AS CAMPO_DET_34,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_35,"
   g_str_Parame = g_str_Parame & "            '7010'                                                                                              AS CAMPO_DET_36,"
   g_str_Parame = g_str_Parame & "            '-'                                                                                                 AS CAMPO_DET_37,"
   g_str_Parame = g_str_Parame & "            'DET2'                                                                                              AS CAMPO_DET2_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_DET2_02," '001-
   g_str_Parame = g_str_Parame & "            '002'                                                                                               AS CAMPO_DET2_03,                " '-- Número de orden de ítem
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_04,                "
   g_str_Parame = g_str_Parame & "            'OTROS IMPORTES ' ||"
   g_str_Parame = g_str_Parame & "            SUBSTR(TO_CHAR(TO_DATE(CAJMOV_FECMOV,'YYYYMMDD'),'MONTH','NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(CAJMOV_FECMOV,1,4)  AS CAMPO_DET2_05,"
   g_str_Parame = g_str_Parame & "            1.000                                                                                               AS CAMPO_DET2_06,"
   g_str_Parame = g_str_Parame & "            'NIU'                                                                                               AS CAMPO_DET2_07,                " '--CATALOGO N°3.-
   g_str_Parame = g_str_Parame & "            (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)"
   'g_str_Parame = g_str_Parame & "            A.CAJMOV_CAPITA + A.CAJMOV_OTRCAR + A.CAJMOV_CAPBBP + A.CAJMOV_INTBBP + A.CAJMOV_INTMOR + A.CAJMOV_INTCOM + A.CAJMOV_GASCOB + A.CAJMOV_OTRGAS"
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_08,"
   g_str_Parame = g_str_Parame & "            1.000 * (A.CAJMOV_IMPPAG - A.CAJMOV_INTERE)"
   'g_str_Parame = g_str_Parame & "            1.000 * ( A.CAJMOV_CAPITA + A.CAJMOV_OTRCAR + A.CAJMOV_CAPBBP + A.CAJMOV_INTBBP + A.CAJMOV_INTMOR + A.CAJMOV_INTCOM + A.CAJMOV_GASCOB + A.CAJMOV_OTRGAS)"
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_09,"
   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_DET2_10_1,              " '--CATALOGO N°5,7 u 8.-
   g_str_Parame = g_str_Parame & "            0.00                                                                                                AS CAMPO_DET2_10_2,"
   g_str_Parame = g_str_Parame & "            '30'                                                                                                AS CAMPO_DET2_10_3,"
   g_str_Parame = g_str_Parame & "            A.CAJMOV_IMPPAG - A.CAJMOV_INTERE "
   'g_str_Parame = g_str_Parame & "            A.CAJMOV_CAPITA + A.CAJMOV_OTRCAR + A.CAJMOV_CAPBBP + A.CAJMOV_INTBBP + A.CAJMOV_INTMOR + A.CAJMOV_INTCOM + A.CAJMOV_GASCOB + A.CAJMOV_OTRGAS"
   g_str_Parame = g_str_Parame & "                                                                                                                AS CAMPO_DET2_11,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_12,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_13,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_14,"
   g_str_Parame = g_str_Parame & "            '84121501'                                                                                          AS CAMPO_DET2_15,                " '--CATALOGO N°15.-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_16,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_17,                " '--NRO CONTRATO O PRESTAMO
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_18,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_19,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_20,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_21,                " '--CATALOGO N°26.- Si es construcción /adquisición
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_22,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_23,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_24,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_25,                " '--CATALOGO N°27.- VERIFICAR
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_26,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_27,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_28,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_29,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_30,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_31,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_32,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_33,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_34,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_35,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_36,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET2_37,"
''   g_str_Parame = g_str_Parame & "            'DET3'                                                                                              AS CAMPO_DET3_01,"
''   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_DET3_02," '001-
''   g_str_Parame = g_str_Parame & "            '003'                                                                                               AS CAMPO_DET3_03,                " '-- Número de orden de ítem
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_04,                "
''   g_str_Parame = g_str_Parame & "            'SEGUROS ' ||"
''   g_str_Parame = g_str_Parame & "            SUBSTR(TO_CHAR(TO_DATE(CAJMOV_FECMOV,'YYYYMMDD'),'MONTH', 'NLS_DATE_LANGUAGE = SPANISH'),1,3) || '-' || SUBSTR(CAJMOV_FECMOV,1,4)  AS CAMPO_DET3_05,"
''   g_str_Parame = g_str_Parame & "            1                                                                                                   AS CAMPO_DET3_06,"
''   g_str_Parame = g_str_Parame & "            'NIU'                                                                                               AS CAMPO_DET3_07,                " '--CATALOGO N°3.-
''   g_str_Parame = g_str_Parame & "            (A.CAJMOV_SEGDES + A.CAJMOV_SEGVIV) - (A.CAJMOV_SEGDES + A.CAJMOV_SEGVIV)*0.18/*A.CAJMOV_SEGDES + A.CAJMOV_SEGVIV*/                                                                   AS CAMPO_DET3_08,"
''   g_str_Parame = g_str_Parame & "            1 * (A.CAJMOV_SEGDES + A.CAJMOV_SEGVIV)                                                             AS CAMPO_DET3_09,"
''   g_str_Parame = g_str_Parame & "            '1000'                                                                                              AS CAMPO_DET3_10_1,              " '--CATALOGO N°5,7 u 8.-
''   g_str_Parame = g_str_Parame & "            (A.CAJMOV_SEGDES + A.CAJMOV_SEGVIV)*0.18                                                            AS CAMPO_DET3_10_2,"
''   g_str_Parame = g_str_Parame & "            '10'                                                                                                AS CAMPO_DET3_10_3,"
''   g_str_Parame = g_str_Parame & "            A.CAJMOV_SEGDES + A.CAJMOV_SEGVIV                                                                   AS CAMPO_DET3_11,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_12,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_13,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_14,"
''   g_str_Parame = g_str_Parame & "            '84121501'                                                                                          AS CAMPO_DET3_15,                " '--CATALOGO N°15.-
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_16,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_17,                " '--NRO CONTRATO O PRESTAMO
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_18,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_19,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_20,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_21,                " '--CATALOGO N°26.- Si es construcción /adquisición
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_22,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_23,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_24,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_25,                " '--CATALOGO N°27.- VERIFICAR
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_26,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_27,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_28,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_29,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_30,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_31,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_32,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_33,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_34,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_35,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_36,"
''   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_DET3_37,"
   g_str_Parame = g_str_Parame & "            'ADI1'                                                                                              AS CAMPO_ADI_01,"
   g_str_Parame = g_str_Parame & "            'F'                                                                                                 AS CAMPO_ADI_02," '001-
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_ADI_03,"
   g_str_Parame = g_str_Parame & "            ''                                                                                                  AS CAMPO_ADI_04,"
   g_str_Parame = g_str_Parame & "            TRIM(A.CAJMOV_NUMOPE)                                                                               AS OPERACION, "
   g_str_Parame = g_str_Parame & "            CAJMOV_NUMMOV                                                                                       AS NUMERO_MOVIMIENTO "
   g_str_Parame = g_str_Parame & "       FROM OPE_CAJMOV A "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES B ON B.PARDES_CODGRP = '204' AND B.PARDES_CODITE = A.CAJMOV_MONPAG "
   g_str_Parame = g_str_Parame & "            INNER JOIN CNTBL_CATSUN C ON C.CATSUN_NROCAT = 2 AND C.CATSUN_DESCRI = TRIM(B.PARDES_DESCRI) "
   g_str_Parame = g_str_Parame & "            INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = A.CAJMOV_TIPDOC AND D.DATGEN_NUMDOC = A.CAJMOV_NUMDOC "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 201 AND E.PARDES_CODITE = D.DATGEN_TIPVIA "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES F ON F.PARDES_CODGRP = 202 AND F.PARDES_CODITE = D.DATGEN_TIPZON "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES G ON G.PARDES_CODGRP = 101 AND G.PARDES_CODITE = D.DATGEN_UBIGEO "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES H ON H.PARDES_CODGRP = 101 AND H.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,2)||'0000' "
   g_str_Parame = g_str_Parame & "            INNER JOIN MNT_PARDES I ON I.PARDES_CODGRP = 101 AND I.PARDES_CODITE = SUBSTR(D.DATGEN_UBIGEO,1,4)||'00' "
   g_str_Parame = g_str_Parame & "            INNER JOIN CRE_HIPMAE J ON J.HIPMAE_NUMOPE = A.CAJMOV_NUMOPE "
   g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_HIPGAR K ON K.HIPGAR_NUMOPE = J.HIPMAE_NUMOPE AND K.HIPGAR_BIEGAR = 1 "
   g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_SOLMAE L ON L.SOLMAE_NUMERO = J.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "             LEFT JOIN CRE_SOLINM P  ON P.SOLINM_NUMSOL = J.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES Q  ON Q.PARDES_CODGRP = 201 AND Q.PARDES_CODITE = P.SOLINM_TIPVIA "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES R  ON R.PARDES_CODGRP = 202 AND R.PARDES_CODITE = P.SOLINM_TIPZON "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES S  ON S.PARDES_CODGRP = 101 AND S.PARDES_CODITE = P.SOLINM_UBIGEO "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES T  ON T.PARDES_CODGRP = 101 AND T.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,4)||'00' "
   g_str_Parame = g_str_Parame & "             LEFT JOIN MNT_PARDES U  ON U.PARDES_CODGRP = 101 AND U.PARDES_CODITE = SUBSTR(P.SOLINM_UBIGEO,1,2)||'0000' "
   g_str_Parame = g_str_Parame & "      WHERE CAJMOV_SUCMOV IS NOT NULL "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_USUMOV IS NOT NULL "
   
   If Chk_FecAct.Value = 0 Then
      g_str_Parame = g_str_Parame & "        AND CAJMOV_FECDEP = '" & l_str_FecCar & "' "
   Else
      g_str_Parame = g_str_Parame & "        AND CAJMOV_FECDEP <= '" & l_str_FecCar & "' "
   End If
   g_str_Parame = g_str_Parame & "        AND CAJMOV_NUMMOV > 0 "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_CODBAN IS NOT NULL "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_FLGPRO = 0 "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_TIPMOV = '1102' "
   g_str_Parame = g_str_Parame & "        AND CAJMOV_TIPDOC IN (1, 6) "                      '1-DNI y 6-RUC
   g_str_Parame = g_str_Parame & "      ORDER BY A.CAJMOV_FECMOV , A.CAJMOV_NUMMOV "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se ejecutó consulta principal, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontró ningún registro.", vbExclamation, modgen_g_str_NomPlt
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro anterior en OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
      Exit Sub
   End If
    
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
      Call fs_Obtener_Codigo("01", r_lng_Contad, r_int_SerFac, r_lng_NumFac)
      
      moddat_g_str_NumOpe = g_rst_Princi!OPERACION
      moddat_g_str_Codigo = g_rst_Princi!NUMERO_MOVIMIENTO
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELE (      "
      g_str_Parame = g_str_Parame & " DOCELE_CODIGO                 , "
      g_str_Parame = g_str_Parame & " DOCELE_NUMOPE                 , "
      g_str_Parame = g_str_Parame & " DOCELE_NUMMOV                 , "
      g_str_Parame = g_str_Parame & " DOCELE_FECPRO                 , "
      g_str_Parame = g_str_Parame & " DOCELE_FECAUT                 , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_FECEMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_HOREMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_TIPMON             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_NUMORC             , "
      g_str_Parame = g_str_Parame & " DOCELE_IDE_FECVCT             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_TIPDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_NUMDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_NOMCOM             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_DENOMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_UBIGEO             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_DIRCOM             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_URBANI             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_PROVIN             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_DEPART             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_DISTRI             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_CODPAI             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_TELEMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_EMI_COREMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_TIPDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_NUMDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_DENOMI             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_DIRCOM             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_DEPART             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_PROVIN             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_DISTRI             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_CODPAI             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_TELREC             , "
      g_str_Parame = g_str_Parame & " DOCELE_REC_CORREC             , "
      g_str_Parame = g_str_Parame & " DOCELE_DRF_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_DRF_TIPDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_DRF_NUMDOC             , "
      g_str_Parame = g_str_Parame & " DOCELE_DRF_CODMOT             , "
      g_str_Parame = g_str_Parame & " DOCELE_DRF_DESMOT             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRV      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRV      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEINA      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEINA      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXO      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXO      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEGRA      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEGRA      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_OPEEXP      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTVTA_OPEEXP      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_PERCEP      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_REGPER      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_BASIMP_PERCEP      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOPER             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOTOT_PERCEP      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIMP             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_MTOIMP             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_OTRCAR             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_CODIGO_TOTDSC      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTDSC             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_IMPTOT_DOCUME      , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_DSCGLO             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_INFPPG             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TOTANT             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_TIPOPE             , "
      g_str_Parame = g_str_Parame & " DOCELE_CAB_LEYEND             , "
      g_str_Parame = g_str_Parame & " DOCELE_ADI_SERNUM             , "
      g_str_Parame = g_str_Parame & " DOCELE_ADI_TITADI             , "
      g_str_Parame = g_str_Parame & " DOCELE_ADI_VALADI             , "
      g_str_Parame = g_str_Parame & " DOCELE_FLGENV                 , "
      g_str_Parame = g_str_Parame & " DOCELE_FLGRPT                 , "
      g_str_Parame = g_str_Parame & " DOCELE_SITUAC                 , "
      g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
      g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
      g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
      g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
      g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
      g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
      g_str_Parame = g_str_Parame & " VALUES ( "
      g_str_Parame = g_str_Parame & "" & r_lng_Contad & " , "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
      g_str_Parame = g_str_Parame & "" & moddat_g_str_Codigo & " , "
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "     , "
      g_str_Parame = g_str_Parame & " NULL, "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_03 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_04 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_05 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_06 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_07 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_IDE_08 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_03 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_04 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_05 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_06 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_07 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_08 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_09 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_10 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_11 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_12 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_13 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_14 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_EMI_15 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_03 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_04 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_05 & "'                   , "
      
      If IsNull(g_rst_Princi!CAMPO_REC_06) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "'" & Mid(Replace(g_rst_Princi!CAMPO_REC_06, "  ", " "), 1, 100) & "'                                    , "
      End If
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_07 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_08 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_09 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_10 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_11 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_REC_12 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_03 & "'                   , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_04 & "'                   , "
      
      If IsNull(g_rst_Princi!CAMPO_DRF_05) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DRF_05 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DRF_06 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_03 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_04) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_04 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_05 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_06) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_06 & "                    , "
      End If
            
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_07 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_08) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_08 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_09 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_10) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_10 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_11 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_12) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_12 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_13 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_14 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_15) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_15 & "                    , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_CAB_16) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_16 & "                    , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_CAB_17) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_17 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_18_1 & "'                   , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_18_2) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_18_2 & "                  , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_CAB_19) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_19 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_20 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_21) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_21 & "                    , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_CAB_22) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_22 & "                    , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_CAB_23) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_23 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_24 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_CAB_25) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_CAB_25 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_26 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_CAB_27 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_03 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_ADI_04 & "'                     , "
      g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
      g_str_Parame = g_str_Parame & "" & 0 & "                                               , "
      g_str_Parame = g_str_Parame & "" & 1 & "                                               , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
      g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
      g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELE, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
         Exit Sub
      End If
      DoEvents: DoEvents: DoEvents

      
      ''INTERES COMPENSATORIO
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
      g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
      g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
      g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
      g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
      g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
      g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
      g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
      g_str_Parame = g_str_Parame & " VALUES (                        "
      g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_03 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_04 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_05 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_DET_06) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_06 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_07 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_DET_08) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_08 & "                    , "
      End If

      If IsNull(g_rst_Princi!CAMPO_DET_09) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_09 & "                    , "
      End If

      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_1 & "'                   , "
      
      If IsNull(g_rst_Princi!CAMPO_DET_10_2) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_10_2 & "                  , "
      End If

      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_10_3 & "'                   , "
      
      If IsNull(g_rst_Princi!CAMPO_DET_11) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_11 & "                    , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_DET_12) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_12 & "                    , "
      End If
   
      If IsNull(g_rst_Princi!CAMPO_DET_13) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_13 & "                    , "
      End If

      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_14 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_15 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_16 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_17 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_18 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_19 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_20 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_DET_21) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_21 & "   , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_22 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_23 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_24 & "'                     , "
      
      If IsNull(g_rst_Princi!CAMPO_DET_25) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET_25 & "                    , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_26 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_27 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_28 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_29 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_30 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_31 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_32 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_33 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_34 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_35 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_36 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET_37 & "'                     , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "'                           , "
      g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
      g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de INTERES COMPENSATORIO, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
         Exit Sub
      End If
      DoEvents: DoEvents: DoEvents
      
                                   
      ''OTROS IMPORTES
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET ("
      g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
      g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
      g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
      g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
      g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
      g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
      g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
      g_str_Parame = g_str_Parame & " VALUES ("
      g_str_Parame = g_str_Parame & "" & r_lng_Contad & "           , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_03 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_04 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_05 & "'                    , "
      If IsNull(g_rst_Princi!CAMPO_DET2_06) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_06 & "                   , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_07 & "'                    , "
      
      If IsNull(g_rst_Princi!CAMPO_DET2_08) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_08 & "                   , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_DET2_09) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_09 & "                   , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_1 & "'                  , "
      
      If IsNull(g_rst_Princi!CAMPO_DET2_10_2) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_10_2 & "                 , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_10_3 & "'                  , "
      
      If IsNull(g_rst_Princi!CAMPO_DET2_11) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_11 & "                   , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_DET2_12) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_12 & "                   , "
      End If
      
      If IsNull(g_rst_Princi!CAMPO_DET2_13) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_13 & "                   , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_14 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_15 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_16 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_17 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_18 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_19 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_20 & "'                    , "
      
      If IsNull(g_rst_Princi!CAMPO_DET2_21) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_21 & "                   , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_22 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_23 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_24 & "'                    , "
      
      If IsNull(g_rst_Princi!CAMPO_DET2_25) Then
         g_str_Parame = g_str_Parame & " NULL, "
      Else
         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET2_25 & "                   , "
      End If
      
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_26 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_27 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_28 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_29 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_30 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_31 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_32 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_33 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_34 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_35 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_36 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET2_37 & "'                    , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
      g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
      g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede insertar en la tabla CNTBL_DOCELEDET, ingreso de OTROS IMPORTES, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
         Exit Sub
      End If
                                                                                                

      DoEvents: DoEvents: DoEvents
      
'''      ''SEGUROS
'''      g_str_Parame = ""
'''      g_str_Parame = g_str_Parame & " INSERT INTO CNTBL_DOCELEDET (   "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_CODIGO              , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_SERNUM          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMITE          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRD          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DESPRD          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CANTID          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_UNIDAD          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALUNI          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PUNVTA          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIMP          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_MTOIMP          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPAFE          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALVTA          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_VALREF          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DSTITE          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_NUMPLA          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODSUN          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODCON          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_NROCON          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_FECOTO   , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_FECOTO          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_TIPPRE   , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_TIPPRE          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PARREG   , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PARREG          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_PRIVIV   , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PRIVIV          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODIGO_DIRCOM   , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DIRCOM          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODUBI          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_UBIGEO          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODURB          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_URBANI          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDPT          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DEPART          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODPRV          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_PROVIN          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_CODDIS          , "
'''      g_str_Parame = g_str_Parame & " DOCELEDET_DET_DISTRI          , "
'''      g_str_Parame = g_str_Parame & " SEGUSUCRE                     , "
'''      g_str_Parame = g_str_Parame & " SEGFECCRE                     , "
'''      g_str_Parame = g_str_Parame & " SEGHORCRE                     , "
'''      g_str_Parame = g_str_Parame & " SEGPLTCRE                     , "
'''      g_str_Parame = g_str_Parame & " SEGTERCRE                     , "
'''      g_str_Parame = g_str_Parame & " SEGSUCCRE                     ) "
'''      g_str_Parame = g_str_Parame & " VALUES ("
'''      g_str_Parame = g_str_Parame & "" & r_lng_Contad & "                                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_02 & Format(r_int_SerFac, "000") & "-" & Format(r_lng_NumFac, "00000000") & "' , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_03 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_04 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_05 & "'                    , "
'''
'''      If IsNull(g_rst_Princi!CAMPO_DET3_06) Then
'''         g_str_Parame = g_str_Parame & " NULL, "
'''      Else
'''         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET3_06 & "                   , "
'''      End If
'''
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_07 & "'                    , "
'''
'''      If IsNull(g_rst_Princi!CAMPO_DET3_08) Then
'''         g_str_Parame = g_str_Parame & " NULL, "
'''      Else
'''         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET3_08 & "                   , "
'''      End If
'''
'''      If IsNull(g_rst_Princi!CAMPO_DET3_09) Then
'''         g_str_Parame = g_str_Parame & " NULL, "
'''      Else
'''         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET3_09 & "                   , "
'''      End If
'''
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_10_1 & "'                  , "
'''
'''      If IsNull(g_rst_Princi!CAMPO_DET3_10_2) Then
'''         g_str_Parame = g_str_Parame & " NULL, "
'''      Else
'''         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET3_10_2 & "                 , "
'''      End If
'''
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_10_3 & "'                  , "
'''
'''      If IsNull(g_rst_Princi!CAMPO_DET3_11) Then
'''         g_str_Parame = g_str_Parame & " NULL, "
'''      Else
'''         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET3_11 & "                   , "
'''      End If
'''
'''      If IsNull(g_rst_Princi!CAMPO_DET3_12) Then
'''         g_str_Parame = g_str_Parame & " NULL, "
'''      Else
'''         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET3_12 & "                   , "
'''      End If
'''
'''      If IsNull(g_rst_Princi!CAMPO_DET3_13) Then
'''         g_str_Parame = g_str_Parame & " NULL, "
'''      Else
'''         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET3_13 & "                   , "
'''      End If
'''
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_14 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_15 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_16 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_17 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_18 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_19 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_20 & "'                    , "
'''
'''      If IsNull(g_rst_Princi!CAMPO_DET3_21) Then
'''         g_str_Parame = g_str_Parame & " NULL, "
'''      Else
'''         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET3_21 & "                   , "
'''      End If
'''
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_22 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_23 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_24 & "'                    , "
'''
'''      If IsNull(g_rst_Princi!CAMPO_DET3_25) Then
'''         g_str_Parame = g_str_Parame & " NULL, "
'''      Else
'''         g_str_Parame = g_str_Parame & "" & g_rst_Princi!CAMPO_DET3_25 & "                   , "
'''      End If
'''
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_26 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_27 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_28 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_29 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_30 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_31 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_32 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_33 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_34 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_35 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_36 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!CAMPO_DET3_37 & "'                    , "
'''      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "'                           , "
'''      g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
'''      g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & "                         , "
'''      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "'                            , "
'''      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "'                           , "
'''      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')                            "
'''
'''      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
'''         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se puede ingresar en la tabla CNTBL_DOCELEDET, ingreso de SEGUROS, Nro Ope:" & moddat_g_str_NumOpe & ", Nro. Mov: " & moddat_g_str_Codigo & ", procedimiento: fs_Genera_FactAnterior")
'''         Exit Sub
'''      End If
'''      DoEvents: DoEvents: DoEvents
      
      'ACTUALIZA EL CAMPO CAJMOV_FLGPRO PARA IDENTIFICAR CUALES SE HAN PROCESADO Y YA SE ENCUENTRAN EN CNTBL_DOCELE
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "UPDATE OPE_CAJMOV SET CAJMOV_FLGPRO = 1 "
      g_str_Parame = g_str_Parame & " WHERE CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND CAJMOV_NUMMOV = '" & moddat_g_str_Codigo & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
         Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar CAJMOV_FLGPRO de la tabla OPE_CAJMOV, procedimiento: fs_Genera_FactAnterior")
         Exit Sub
      End If
         
      g_rst_Princi.MoveNext
   Loop
      
   Exit Sub
   
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Genera_FactAnterior")

End Sub
Private Sub fs_Generar_Archivo_01_03(ByVal p_DiaAct As String, ByVal p_TipDoc As String)

Dim r_str_NomRes     As String
Dim r_int_NumRes     As Integer
Dim r_str_Nombre     As String
Dim r_str_MtoLtr     As String
Dim r_str_Parame     As String

On Error GoTo MyError

   Screen.MousePointer = 11
   
   '*** GENERANDO REPORTE
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DOCELE_CODIGO            , DOCELE_FECPRO            , DOCELE_FECAUT            , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_SERNUM          CAMPO_IDE_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_FECEMI          CAMPO_IDE_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_HOREMI          CAMPO_IDE_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_TIPDOC          CAMPO_IDE_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_TIPMON          CAMPO_IDE_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_NUMORC          CAMPO_IDE_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_FECVCT          CAMPO_IDE_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_SERNUM          CAMPO_EMI_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_TIPDOC          CAMPO_EMI_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_NUMDOC          CAMPO_EMI_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_NOMCOM          CAMPO_EMI_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DENOMI          CAMPO_EMI_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_UBIGEO          CAMPO_EMI_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DIRCOM          CAMPO_EMI_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_URBANI          CAMPO_EMI_09      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_PROVIN          CAMPO_EMI_10      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DEPART          CAMPO_EMI_11      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DISTRI          CAMPO_EMI_12      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_CODPAI          CAMPO_EMI_13      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_TELEMI          CAMPO_EMI_14      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_COREMI          CAMPO_EMI_15      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_SERNUM          CAMPO_REC_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_TIPDOC          CAMPO_REC_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_NUMDOC          CAMPO_REC_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DENOMI          CAMPO_REC_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DIRCOM          CAMPO_REC_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DEPART          CAMPO_REC_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_PROVIN          CAMPO_REC_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DISTRI          CAMPO_REC_09      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_CODPAI          CAMPO_REC_10      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_TELREC          CAMPO_REC_11      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_CORREC          CAMPO_REC_12      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_SERNUM          CAMPO_DRF_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_TIPDOC          CAMPO_DRF_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_NUMDOC          CAMPO_DRF_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_CODMOT          CAMPO_DRF_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_DESMOT          CAMPO_DRF_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_SERNUM          CAMPO_CAB_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEGRV   CAMPO_CAB_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEGRV   CAMPO_CAB_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEINA   CAMPO_CAB_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEINA   CAMPO_CAB_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEEXO   CAMPO_CAB_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEEXO   CAMPO_CAB_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEGRA   CAMPO_CAB_09      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEGRA   CAMPO_CAB_10      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEEXP   CAMPO_CAB_11      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEEXP   CAMPO_CAB_12      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_PERCEP   CAMPO_CAB_13      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_REGPER   CAMPO_CAB_14      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_BASIMP_PERCEP   CAMPO_CAB_15      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_MTOPER          CAMPO_CAB_16      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_MTOTOT_PERCEP   CAMPO_CAB_17      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIMP          CAMPO_CAB_18_1    , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_MTOIMP          CAMPO_CAB_18_2    , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_OTRCAR          CAMPO_CAB_19      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_TOTDSC   CAMPO_CAB_20      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTDSC          CAMPO_CAB_21      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_IMPTOT_DOCUME   CAMPO_CAB_22      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_DSCGLO          CAMPO_CAB_23      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_INFPPG          CAMPO_CAB_24      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTANT          CAMPO_CAB_25      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TIPOPE          CAMPO_CAB_26      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_LEYEND          CAMPO_CAB_27      , "
   g_str_Parame = g_str_Parame & "        DOCELE_ADI_SERNUM          CAMPO_ADI_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_ADI_TITADI          CAMPO_ADI_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_ADI_VALADI          CAMPO_ADI_04        "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_DOCELE "
   g_str_Parame = g_str_Parame & "  WHERE DOCELE_FLGENV = 0 "
   g_str_Parame = g_str_Parame & "    AND DOCELE_IDE_TIPDOC = '" & p_TipDoc & "'"
   g_str_Parame = g_str_Parame & "    AND DOCELE_SITUAC = 1 "
   
   If cmb_TipPro.ItemData(cmb_TipPro.ListIndex) = 9 Then
      g_str_Parame = g_str_Parame & " AND DOCELE_TIPPRO IS NULL " 'not
   Else
      g_str_Parame = g_str_Parame & " AND DOCELE_TIPPRO IS NOT NULL " ' & cmb_TipPro.ItemData(cmb_TipPro.ListIndex) & ""
   End If
   
   g_str_Parame = g_str_Parame & "  ORDER BY DOCELE_CODIGO ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error en la consulta principal de la tabla CNTBL_DOCELE, procedimiento: fs_Generar_Archivo")
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro en la consulta principal de la tabla CNTBL_DOCELE, procedimiento: fs_Generar_Archivo")
      Exit Sub
   End If
   
   'Creando Archivo
   r_str_Nombre = "20511904162-" & p_TipDoc & "-" & p_DiaAct & ".txt"
   r_str_NomRes = l_str_RutFacEnt & "20511904162-" & p_TipDoc & "-" & p_DiaAct & ".txt"
   r_int_NumRes = FreeFile
   
   Open r_str_NomRes For Output As r_int_NumRes
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
         Print #r_int_NumRes, "IDE"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_02), "", Trim(g_rst_Princi!CAMPO_IDE_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_03), "", Trim(g_rst_Princi!CAMPO_IDE_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_04), "", Trim(g_rst_Princi!CAMPO_IDE_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_05), "", Trim(g_rst_Princi!CAMPO_IDE_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_06), "", Trim(g_rst_Princi!CAMPO_IDE_06)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_07), "", Trim(g_rst_Princi!CAMPO_IDE_07)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_08), "", Trim(g_rst_Princi!CAMPO_IDE_08)) & vbCrLf;
                                      
         Print #r_int_NumRes, "EMI"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_02), "", Trim(g_rst_Princi!CAMPO_EMI_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_03), "", Trim(g_rst_Princi!CAMPO_EMI_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_04), "", Trim(g_rst_Princi!CAMPO_EMI_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_05), "", Trim(g_rst_Princi!CAMPO_EMI_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_06), "", Trim(g_rst_Princi!CAMPO_EMI_06)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_07), "", Trim(g_rst_Princi!CAMPO_EMI_07)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_08), "", Trim(g_rst_Princi!CAMPO_EMI_08)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_09), "", Trim(g_rst_Princi!CAMPO_EMI_09)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_10), "", Trim(g_rst_Princi!CAMPO_EMI_10)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_11), "", Trim(g_rst_Princi!CAMPO_EMI_11)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_12), "", Trim(g_rst_Princi!CAMPO_EMI_12)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_13), "", Trim(g_rst_Princi!CAMPO_EMI_13)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_14), "", Trim(g_rst_Princi!CAMPO_EMI_14)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_15), "", Trim(g_rst_Princi!CAMPO_EMI_15)) & vbCrLf;
                   
         
         Print #r_int_NumRes, "REC"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_02), "", Trim(g_rst_Princi!CAMPO_REC_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_03), "", Trim(g_rst_Princi!CAMPO_REC_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_04), "", Trim(g_rst_Princi!CAMPO_REC_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_05), "", Trim(g_rst_Princi!CAMPO_REC_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_06), "", ""); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_07), "", Trim(g_rst_Princi!CAMPO_REC_07)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_08), "", Trim(g_rst_Princi!CAMPO_REC_08)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_09), "", Trim(g_rst_Princi!CAMPO_REC_09)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_10), "", Trim(g_rst_Princi!CAMPO_REC_10)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_11), "", Trim(g_rst_Princi!CAMPO_REC_11)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_12), "", Trim(g_rst_Princi!CAMPO_REC_12)) & vbCrLf;
                   'Trim(g_rst_Princi!CAMPO_REC_06)
                 
         If Not IsNull(g_rst_Princi!CAMPO_DRF_03) And g_rst_Princi!CAMPO_DRF_03 <> "" Then
            Print #r_int_NumRes, "DRF"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_02), "", Trim(g_rst_Princi!CAMPO_DRF_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_03), "", Trim(g_rst_Princi!CAMPO_DRF_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_04), "", Trim(g_rst_Princi!CAMPO_DRF_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_05), "", Trim(g_rst_Princi!CAMPO_DRF_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_06), "", Trim(g_rst_Princi!CAMPO_DRF_06)) & vbCrLf;
                   
         End If
         
         r_str_MtoLtr = fs_NroEnLetras(g_rst_Princi!CAMPO_CAB_22)
         
         Print #r_int_NumRes, "CAB"; "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_02), "", Trim(g_rst_Princi!CAMPO_CAB_02)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_03), "", Trim(g_rst_Princi!CAMPO_CAB_03)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_04), "", Format(Trim(g_rst_Princi!CAMPO_CAB_04), "0.00")); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_05), "", Trim(g_rst_Princi!CAMPO_CAB_05)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_06), "", Trim(g_rst_Princi!CAMPO_CAB_06)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_07), "", Trim(g_rst_Princi!CAMPO_CAB_07)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_08), "", Format(Trim(g_rst_Princi!CAMPO_CAB_08), "0.00")); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_09), "", Trim(g_rst_Princi!CAMPO_CAB_09)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_10), "", Format(Trim(g_rst_Princi!CAMPO_CAB_10), "0.00")); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_11), "", Trim(g_rst_Princi!CAMPO_CAB_11)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_12), "", Format(Trim(g_rst_Princi!CAMPO_CAB_12), "0.00")); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_13), "", Trim(g_rst_Princi!CAMPO_CAB_13)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_14), "", Trim(g_rst_Princi!CAMPO_CAB_14)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_15), "", Trim(g_rst_Princi!CAMPO_CAB_15)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_16), "", Trim(g_rst_Princi!CAMPO_CAB_16)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_17), "", Trim(g_rst_Princi!CAMPO_CAB_17)); "|"; _
                  "[" & IIf(IsNull(g_rst_Princi!CAMPO_CAB_18_1), "", Trim(g_rst_Princi!CAMPO_CAB_18_1)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_18_2), "", Format(Trim(g_rst_Princi!CAMPO_CAB_18_2), "0.00")); "]|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_19), "", Trim(g_rst_Princi!CAMPO_CAB_19)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_20), "", Trim(g_rst_Princi!CAMPO_CAB_20)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_21), "", Trim(g_rst_Princi!CAMPO_CAB_21)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_22), "", Trim(g_rst_Princi!CAMPO_CAB_22)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_23), "", Trim(g_rst_Princi!CAMPO_CAB_23)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_24), "", Trim(g_rst_Princi!CAMPO_CAB_24)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_25), "", Trim(g_rst_Princi!CAMPO_CAB_25)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_26), "", Trim(g_rst_Princi!CAMPO_CAB_26)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_27), "", Trim(g_rst_Princi!CAMPO_CAB_27) & "|" & r_str_MtoLtr & "]") & vbCrLf;
                  
                  
                  
         'Generar detalle de las facturas
         moddat_g_str_CodGen = g_rst_Princi!DOCELE_CODIGO
            
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " SELECT DOCELEDET_CODIGO                              , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_SERNUM          CAMPO_DET_02    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_NUMITE          CAMPO_DET_03    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODPRD          CAMPO_DET_04    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DESPRD          CAMPO_DET_05    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CANTID          CAMPO_DET_06    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_UNIDAD          CAMPO_DET_07    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_VALUNI          CAMPO_DET_08    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PUNVTA          CAMPO_DET_09    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIMP          CAMPO_DET_10_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_MTOIMP          CAMPO_DET_10_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_TIPAFE          CAMPO_DET_10_3  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_VALVTA          CAMPO_DET_11    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_VALREF          CAMPO_DET_12    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DSTITE          CAMPO_DET_13    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_NUMPLA          CAMPO_DET_14    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODSUN          CAMPO_DET_15    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODCON          CAMPO_DET_16    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_NROCON          CAMPO_DET_17    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_FECOTO   CAMPO_DET_18    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_FECOTO          CAMPO_DET_19    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_TIPPRE   CAMPO_DET_20    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_TIPPRE          CAMPO_DET_21    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_PARREG   CAMPO_DET_22    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PARREG          CAMPO_DET_23    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_PRIVIV   CAMPO_DET_24    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PRIVIV          CAMPO_DET_25    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_DIRCOM   CAMPO_DET_26    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DIRCOM          CAMPO_DET_27    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODUBI          CAMPO_DET_28    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_UBIGEO          CAMPO_DET_29    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODURB          CAMPO_DET_30    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_URBANI          CAMPO_DET_31    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODDPT          CAMPO_DET_32    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DEPART          CAMPO_DET_33    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODPRV          CAMPO_DET_34    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PROVIN          CAMPO_DET_35    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODDIS          CAMPO_DET_36    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DISTRI          CAMPO_DET_37      "
         g_str_Parame = g_str_Parame & "   FROM CNTBL_DOCELEDET "
         g_str_Parame = g_str_Parame & "  WHERE DOCELEDET_CODIGO = " & CLng(moddat_g_str_CodGen)
         g_str_Parame = g_str_Parame & "  ORDER BY DOCELEDET_CODIGO, DOCELEDET_DET_NUMITE ASC "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se lee la consulta en la tabla CNTBL_DOCELEDET, Nro Ope:" & moddat_g_str_CodGen & ", procedimiento: fs_Generar_Archivo")
            Exit Sub
         End If
              
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            g_rst_Genera.Close
            Set g_rst_Genera = Nothing
            Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No hay ningun registro en la tabla CNTBL_DOCELEDET, Nro Ope:" & moddat_g_str_CodGen & ", procedimiento: fs_Generar_Archivo")
            Exit Sub
         End If
         
         g_rst_Genera.MoveFirst
         
         Do While Not g_rst_Genera.EOF
            
            
            Print #r_int_NumRes, "DET"; "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_02), "", Trim(g_rst_Genera!CAMPO_DET_02)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_03), "", Trim(g_rst_Genera!CAMPO_DET_03)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_04), "", Trim(g_rst_Genera!CAMPO_DET_04)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_05), "", Trim(g_rst_Genera!CAMPO_DET_05)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_06), "", Format(Trim(g_rst_Genera!CAMPO_DET_06), "0.000")); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_07), "", Trim(g_rst_Genera!CAMPO_DET_07)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_08), "", Trim(g_rst_Genera!CAMPO_DET_08)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_09), "", Trim(g_rst_Genera!CAMPO_DET_09)); "|"; _
                  "[" & IIf(IsNull(g_rst_Genera!CAMPO_DET_10_1), "", Trim(g_rst_Genera!CAMPO_DET_10_1)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_10_2), "", Format(Trim(g_rst_Genera!CAMPO_DET_10_2), "0.00")); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_10_3), "", Trim(g_rst_Genera!CAMPO_DET_10_3)); "]|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_11), "", Format(Trim(g_rst_Genera!CAMPO_DET_11), "0.00")); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_12), "", Trim(g_rst_Genera!CAMPO_DET_12)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_13), "", Trim(g_rst_Genera!CAMPO_DET_13)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_14), "", Trim(g_rst_Genera!CAMPO_DET_14)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15), "", Trim(g_rst_Genera!CAMPO_DET_15)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_16), "", Trim(g_rst_Genera!CAMPO_DET_16)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_17), "", Trim(g_rst_Genera!CAMPO_DET_17)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_18), "", Trim(g_rst_Genera!CAMPO_DET_18)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_19), "", Trim(g_rst_Genera!CAMPO_DET_19)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_20), "", Trim(g_rst_Genera!CAMPO_DET_20)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_21), "", Trim(g_rst_Genera!CAMPO_DET_21)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_22), "", Trim(g_rst_Genera!CAMPO_DET_22)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_23), "", Trim(g_rst_Genera!CAMPO_DET_23)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_24), "", Trim(g_rst_Genera!CAMPO_DET_24)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_25), "", Trim(g_rst_Genera!CAMPO_DET_25)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_26), "", Trim(g_rst_Genera!CAMPO_DET_26)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_27), "", Trim(g_rst_Genera!CAMPO_DET_27)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_28), "", Trim(g_rst_Genera!CAMPO_DET_28)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_29), "", Trim(g_rst_Genera!CAMPO_DET_29)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_30), "", Trim(g_rst_Genera!CAMPO_DET_30)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_31), "", Trim(g_rst_Genera!CAMPO_DET_31)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_32), "", Trim(g_rst_Genera!CAMPO_DET_32)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_33), "", Trim(g_rst_Genera!CAMPO_DET_33)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_34), "", Trim(g_rst_Genera!CAMPO_DET_34)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_35), "", Trim(g_rst_Genera!CAMPO_DET_35)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_36), "", Trim(g_rst_Genera!CAMPO_DET_36)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_37), "", Trim(g_rst_Genera!CAMPO_DET_37)); vbCrLf;
                  
            g_rst_Genera.MoveNext
         Loop
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         Print #r_int_NumRes, "##############################" & vbCrLf; 'vbCr
                     
      g_rst_Princi.MoveNext
      
      DoEvents
         
   Loop
   
               
'   If Not IsNull(g_rst_Princi!CAMPO_ADI_03) And g_rst_Princi!CAMPO_ADI_03 <> "" Then
'      Print #1, "ADI"; "|"; _
'             IIf(IsNull(g_rst_Princi!CAMPO_ADI_02), "", Trim(g_rst_Princi!CAMPO_ADI_02)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_ADI_03), "", Trim(g_rst_Princi!CAMPO_ADI_03)); "|"; _
'             IIf(IsNull(g_rst_Princi!CAMPO_ADI_04), "", Trim(g_rst_Princi!CAMPO_ADI_04)); "|";
'   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Close #1
   
   'Convertir a UTF8 sin BOM
   l_str_RutaArc = r_str_NomRes
   Call fs_Convertir_Utf8NoBom(r_str_NomRes)
   
   'Enviar el archivo en el ftp
   If fs_Cargar_Archivo(r_str_NomRes, r_str_Nombre) = True Then
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   Se envió correctamente el archivo :" & r_str_Nombre & ", procedimiento: fs_Generar_Archivo")
      
      'Leer el archivo enviado para determinar si se envió al Sftp
      Call fs_Leer_Archivo_EnvSFTP(r_str_NomRes)
      
   Else
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se cargó el archivo en SFTP, procedimiento: fs_Generar_Archivo")
   End If
   
   Exit Sub
   
MyError:

   Close #1
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Generar_Archivo")
   
End Sub

Private Sub fs_Generar_Archivo_01_03_NUEVO_FORMATO(ByVal p_DiaAct As String, ByVal p_TipDoc As String)

Dim r_str_NomRes     As String
Dim r_int_NumRes     As Integer
Dim r_str_Nombre     As String
Dim r_str_MtoLtr     As String
Dim r_str_Parame     As String

On Error GoTo MyError

   Screen.MousePointer = 11
   
   '*** GENERANDO REPORTE
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DOCELE_CODIGO            , DOCELE_FECPRO            , DOCELE_FECAUT            , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_SERNUM          CAMPO_IDE_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_FECEMI          CAMPO_IDE_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_HOREMI          CAMPO_IDE_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_TIPDOC          CAMPO_IDE_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_TIPMON          CAMPO_IDE_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_NUMORC          CAMPO_IDE_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_FECVCT          CAMPO_IDE_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_SERNUM          CAMPO_EMI_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_TIPDOC          CAMPO_EMI_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_NUMDOC          CAMPO_EMI_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_NOMCOM          CAMPO_EMI_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DENOMI          CAMPO_EMI_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_UBIGEO          CAMPO_EMI_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DIRCOM          CAMPO_EMI_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_URBANI          CAMPO_EMI_09      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_PROVIN          CAMPO_EMI_10      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DEPART          CAMPO_EMI_11      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DISTRI          CAMPO_EMI_12      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_CODPAI          CAMPO_EMI_13      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_TELEMI          CAMPO_EMI_14      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_COREMI          CAMPO_EMI_15      , "
   
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_CODSUN          CAMPO_EMI_16      , "
   
   g_str_Parame = g_str_Parame & "        DOCELE_REC_SERNUM          CAMPO_REC_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_TIPDOC          CAMPO_REC_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_NUMDOC          CAMPO_REC_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DENOMI          CAMPO_REC_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DIRCOM          CAMPO_REC_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DEPART          CAMPO_REC_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_PROVIN          CAMPO_REC_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DISTRI          CAMPO_REC_09      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_CODPAI          CAMPO_REC_10      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_TELREC          CAMPO_REC_11      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_CORREC          CAMPO_REC_12      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_SERNUM          CAMPO_DRF_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_TIPDOC          CAMPO_DRF_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_NUMDOC          CAMPO_DRF_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_CODMOT          CAMPO_DRF_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_DESMOT          CAMPO_DRF_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_SERNUM          CAMPO_CAB_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEGRV   CAMPO_CAB_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEGRV   CAMPO_CAB_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEINA   CAMPO_CAB_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEINA   CAMPO_CAB_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEEXO   CAMPO_CAB_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEEXO   CAMPO_CAB_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEGRA   CAMPO_CAB_09      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEGRA   CAMPO_CAB_10      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEEXP   CAMPO_CAB_11      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEEXP   CAMPO_CAB_12      , "
   
   
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_INDICA_CARDSC   CAMPO_CAB_13      , "
   
'   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_PERCEP   CAMPO_CAB_13      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_REGPER   CAMPO_CAB_14      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_BASIMP_PERCEP   CAMPO_CAB_15      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_MTOPER          CAMPO_CAB_16      , "
'   g_str_Parame = g_str_Parame & "        DOCELE_CAB_MTOTOT_PERCEP   CAMPO_CAB_17      , "
   
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_FACTOR_TASPER   CAMPO_CAB_17      , "

   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIMP          CAMPO_CAB_18_1    , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_MTOIMP          CAMPO_CAB_18_2    , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_OTRCAR          CAMPO_CAB_19      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_TOTDSC   CAMPO_CAB_20      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTDSC          CAMPO_CAB_21      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_IMPTOT_DOCUME   CAMPO_CAB_22      , "
   
'   g_str_Parame = g_str_Parame & "        DOCELE_CAB_DSCGLO          CAMPO_CAB_23      , "

   g_str_Parame = g_str_Parame & "        DOCELE_CAB_INFPPG          CAMPO_CAB_23      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTANT          CAMPO_CAB_24      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TIPOPE          CAMPO_CAB_25      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_LEYEND          CAMPO_CAB_26      , "
   
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_MTOTOT_IMP      CAMPO_CAB_27      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CARDSC          CAMPO_CAB_28_1    , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODMOT_CARDSC   CAMPO_CAB_28_2    , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_FACTOR_CARDSC   CAMPO_CAB_28_3    , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_MONTO_CARDSC    CAMPO_CAB_28_4    , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_MONTO_CARDSC    CAMPO_CAB_28_5    , "
   
   g_str_Parame = g_str_Parame & "        DOCELE_ADI_SERNUM          CAMPO_ADI_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_ADI_TITADI          CAMPO_ADI_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_ADI_VALADI          CAMPO_ADI_04        "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_DOCELE "
   g_str_Parame = g_str_Parame & "  WHERE DOCELE_FLGENV = 0 "
   g_str_Parame = g_str_Parame & "    AND DOCELE_IDE_TIPDOC = '" & p_TipDoc & "'"
   g_str_Parame = g_str_Parame & "    AND DOCELE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "  ORDER BY DOCELE_CODIGO ASC "

    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error en la consulta principal de la tabla CNTBL_DOCELE, procedimiento: fs_Generar_Archivo")
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro en la consulta principal de la tabla CNTBL_DOCELE, procedimiento: fs_Generar_Archivo")
      Exit Sub
   End If
   
   'Creando Archivo
   r_str_Nombre = "20511904162-" & p_TipDoc & "-" & p_DiaAct & ".txt"
   r_str_NomRes = l_str_RutFacEnt & "20511904162-" & p_TipDoc & "-" & p_DiaAct & ".txt"
   r_int_NumRes = FreeFile
   
   Open r_str_NomRes For Output As r_int_NumRes
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
         Print #r_int_NumRes, "IDE"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_02), "", Trim(g_rst_Princi!CAMPO_IDE_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_03), "", Trim(g_rst_Princi!CAMPO_IDE_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_04), "", Trim(g_rst_Princi!CAMPO_IDE_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_05), "", Trim(g_rst_Princi!CAMPO_IDE_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_06), "", Trim(g_rst_Princi!CAMPO_IDE_06)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_07), "", Trim(g_rst_Princi!CAMPO_IDE_07)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_08), "", Trim(g_rst_Princi!CAMPO_IDE_08)) & vbCrLf;
                                      
         Print #r_int_NumRes, "EMI"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_02), "", Trim(g_rst_Princi!CAMPO_EMI_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_03), "", Trim(g_rst_Princi!CAMPO_EMI_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_04), "", Trim(g_rst_Princi!CAMPO_EMI_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_05), "", Trim(g_rst_Princi!CAMPO_EMI_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_06), "", Trim(g_rst_Princi!CAMPO_EMI_06)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_07), "", Trim(g_rst_Princi!CAMPO_EMI_07)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_08), "", Trim(g_rst_Princi!CAMPO_EMI_08)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_09), "", Trim(g_rst_Princi!CAMPO_EMI_09)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_10), "", Trim(g_rst_Princi!CAMPO_EMI_10)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_11), "", Trim(g_rst_Princi!CAMPO_EMI_11)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_12), "", Trim(g_rst_Princi!CAMPO_EMI_12)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_13), "", Trim(g_rst_Princi!CAMPO_EMI_13)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_14), "", Trim(g_rst_Princi!CAMPO_EMI_14)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_15), "", Trim(g_rst_Princi!CAMPO_EMI_15)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_16), "", Trim(g_rst_Princi!CAMPO_EMI_16)) & vbCrLf;
                   
         
         Print #r_int_NumRes, "REC"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_02), "", Trim(g_rst_Princi!CAMPO_REC_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_03), "", Trim(g_rst_Princi!CAMPO_REC_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_04), "", Trim(g_rst_Princi!CAMPO_REC_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_05), "", Trim(g_rst_Princi!CAMPO_REC_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_06), "", Trim(g_rst_Princi!CAMPO_REC_06)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_07), "", Trim(g_rst_Princi!CAMPO_REC_07)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_08), "", Trim(g_rst_Princi!CAMPO_REC_08)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_09), "", Trim(g_rst_Princi!CAMPO_REC_09)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_10), "", Trim(g_rst_Princi!CAMPO_REC_10)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_11), "", Trim(g_rst_Princi!CAMPO_REC_11)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_12), "", Trim(g_rst_Princi!CAMPO_REC_12)) & vbCrLf;
                   
                   
         If Not IsNull(g_rst_Princi!CAMPO_DRF_03) And g_rst_Princi!CAMPO_DRF_03 <> "" Then
            Print #r_int_NumRes, "DRF"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_02), "", Trim(g_rst_Princi!CAMPO_DRF_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_03), "", Trim(g_rst_Princi!CAMPO_DRF_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_04), "", Trim(g_rst_Princi!CAMPO_DRF_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_05), "", Trim(g_rst_Princi!CAMPO_DRF_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_06), "", Trim(g_rst_Princi!CAMPO_DRF_06)) & vbCrLf;
                   
         End If
         
         r_str_MtoLtr = fs_NroEnLetras(g_rst_Princi!CAMPO_CAB_22)
         
         Print #r_int_NumRes, "CAB"; "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_02), "", Trim(g_rst_Princi!CAMPO_CAB_02)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_03), "", Trim(g_rst_Princi!CAMPO_CAB_03)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_04), "", Format(Trim(g_rst_Princi!CAMPO_CAB_04), "0.00")); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_05), "", Trim(g_rst_Princi!CAMPO_CAB_05)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_06), "", Trim(g_rst_Princi!CAMPO_CAB_06)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_07), "", Trim(g_rst_Princi!CAMPO_CAB_07)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_08), "", Format(Trim(g_rst_Princi!CAMPO_CAB_08), "0.00")); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_09), "", Trim(g_rst_Princi!CAMPO_CAB_09)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_10), "", Format(Trim(g_rst_Princi!CAMPO_CAB_10), "0.00")); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_11), "", Trim(g_rst_Princi!CAMPO_CAB_11)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_12), "", Format(Trim(g_rst_Princi!CAMPO_CAB_12), "0.00")); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_13), "", Trim(g_rst_Princi!CAMPO_CAB_13)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_14), "", Trim(g_rst_Princi!CAMPO_CAB_14)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_15), "", Trim(g_rst_Princi!CAMPO_CAB_15)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_16), "", Trim(g_rst_Princi!CAMPO_CAB_16)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_17), "", Trim(g_rst_Princi!CAMPO_CAB_17)); "|"; _
                  "[" & IIf(IsNull(g_rst_Princi!CAMPO_CAB_18_1), "", Trim(g_rst_Princi!CAMPO_CAB_18_1)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_18_2), "", Format(Trim(g_rst_Princi!CAMPO_CAB_18_2), "0.00")); "]|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_19), "", Trim(g_rst_Princi!CAMPO_CAB_19)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_20), "", Trim(g_rst_Princi!CAMPO_CAB_20)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_21), "", Trim(g_rst_Princi!CAMPO_CAB_21)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_22), "", Trim(g_rst_Princi!CAMPO_CAB_22)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_23), "", Trim(g_rst_Princi!CAMPO_CAB_23)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_24), "", Trim(g_rst_Princi!CAMPO_CAB_24)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_25), "", Trim(g_rst_Princi!CAMPO_CAB_25)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_26), "", Trim(g_rst_Princi!CAMPO_CAB_26)); "|" & r_str_MtoLtr & "]"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_27), "", Trim(g_rst_Princi!CAMPO_CAB_27)); "|"; _
                  "[" & IIf(IsNull(g_rst_Princi!CAMPO_CAB_28_1), "", Trim(g_rst_Princi!CAMPO_CAB_28_1)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_28_2), "", Trim(g_rst_Princi!CAMPO_CAB_28_2)); IIf(IsNull(g_rst_Princi!CAMPO_CAB_28_3), "", Format(Trim(g_rst_Princi!CAMPO_CAB_28_3), "0.00")); IIf(IsNull(g_rst_Princi!CAMPO_CAB_28_4), "", Format(Trim(g_rst_Princi!CAMPO_CAB_28_4), "0.00")); IIf(IsNull(g_rst_Princi!CAMPO_CAB_28_5), "", Format(Trim(g_rst_Princi!CAMPO_CAB_28_5), "0.00")); "]|" & vbCrLf;
                  
                  'IIf(IsNull(g_rst_Princi!CAMPO_CAB_28), "", Trim(g_rst_Princi!CAMPO_CAB_28) & "|") & vbCrLf;
                  'IIf(IsNull(g_rst_Princi!CAMPO_CAB_23), "", Trim(g_rst_Princi!CAMPO_CAB_23)); "|"; _

         'Generar detalle de las facturas
         moddat_g_str_CodGen = g_rst_Princi!DOCELE_CODIGO
            
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " SELECT DOCELEDET_CODIGO                                , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_SERNUM          CAMPO_DET_02      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_NUMITE          CAMPO_DET_03      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODPRD          CAMPO_DET_04      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DESPRD          CAMPO_DET_05      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CANTID          CAMPO_DET_06      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_UNIDAD          CAMPO_DET_07      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_VALUNI          CAMPO_DET_08      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PUNVTA          CAMPO_DET_09      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIMP          CAMPO_DET_10_1    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_MTOIMP          CAMPO_DET_10_2    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_TIPAFE          CAMPO_DET_10_3    , "
         
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_MTOBAS_IMP      CAMPO_DET_10_4    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_TASTRI          CAMPO_DET_10_5    , "
         
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_VALVTA          CAMPO_DET_11      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_VALREF          CAMPO_DET_12      , "
'         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DSTITE          CAMPO_DET_13     , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_NUMPLA          CAMPO_DET_13      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODSUN          CAMPO_DET_14      , "
         
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_TIPPRE   CAMPO_DET_15_1_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_TIPPRE          CAMPO_DET_15_1_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_PRIVIV   CAMPO_DET_15_2_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PRIVIV          CAMPO_DET_15_2_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_PARREG   CAMPO_DET_15_3_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PARREG          CAMPO_DET_15_3_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODCON          CAMPO_DET_15_4_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_NROCON          CAMPO_DET_15_4_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_FECOTO   CAMPO_DET_15_5_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_FECOTO          CAMPO_DET_15_5_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODUBI          CAMPO_DET_15_6_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_UBIGEO          CAMPO_DET_15_6_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_DIRCOM   CAMPO_DET_15_7_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DIRCOM          CAMPO_DET_15_7_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODURB          CAMPO_DET_15_8_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_URBANI          CAMPO_DET_15_8_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODPRV          CAMPO_DET_15_9_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PROVIN          CAMPO_DET_15_9_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODDIS          CAMPO_DET_15_10_1 , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DISTRI          CAMPO_DET_15_10_2 , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODDPT          CAMPO_DET_15_11_1 , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DEPART          CAMPO_DET_15_11_2 , "
         
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODPRD_GS1      CAMPO_DET_37      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_MTOTOT_IMP      CAMPO_DET_38      , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_INDCAR          CAMPO_DET_39_1    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODCAR          CAMPO_DET_39_2    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_FACCAR          CAMPO_DET_39_3    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DESITE          CAMPO_DET_39_4    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_MTOBAS_CAR      CAMPO_DET_39_5      "
         
         g_str_Parame = g_str_Parame & "   FROM CNTBL_DOCELEDET "
         g_str_Parame = g_str_Parame & "  WHERE DOCELEDET_CODIGO = " & CLng(moddat_g_str_CodGen)
         g_str_Parame = g_str_Parame & "  ORDER BY DOCELEDET_CODIGO, DOCELEDET_DET_NUMITE ASC "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se lee la consulta en la tabla CNTBL_DOCELEDET, Nro Ope:" & moddat_g_str_CodGen & ", procedimiento: fs_Generar_Archivo")
            Exit Sub
         End If
              
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            g_rst_Genera.Close
            Set g_rst_Genera = Nothing
            Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No hay ningun registro en la tabla CNTBL_DOCELEDET, Nro Ope:" & moddat_g_str_CodGen & ", procedimiento: fs_Generar_Archivo")
            Exit Sub
         End If
         
         g_rst_Genera.MoveFirst
         
         Do While Not g_rst_Genera.EOF
         
            Print #r_int_NumRes, "DET"; "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_02), "", Trim(g_rst_Genera!CAMPO_DET_02)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_03), "", Trim(g_rst_Genera!CAMPO_DET_03)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_04), "", Trim(g_rst_Genera!CAMPO_DET_04)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_05), "", Trim(g_rst_Genera!CAMPO_DET_05)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_06), "", Format(Trim(g_rst_Genera!CAMPO_DET_06), "0.000")); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_07), "", Trim(g_rst_Genera!CAMPO_DET_07)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_08), "", Trim(g_rst_Genera!CAMPO_DET_08)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_09), "", Trim(g_rst_Genera!CAMPO_DET_09)); "|"; _
                  "[" & IIf(IsNull(g_rst_Genera!CAMPO_DET_10_1), "", Trim(g_rst_Genera!CAMPO_DET_10_1)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_10_2), "", Format(Trim(g_rst_Genera!CAMPO_DET_10_2), "0.00")); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_10_3), "", Trim(g_rst_Genera!CAMPO_DET_10_3)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_10_4), "", Format(Trim(g_rst_Genera!CAMPO_DET_10_4), "0.00")); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_10_5), "", Format(Trim(g_rst_Genera!CAMPO_DET_10_5), "0.00")); "]|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_11), "", Format(Trim(g_rst_Genera!CAMPO_DET_11), "0.00")); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_12), "", Trim(g_rst_Genera!CAMPO_DET_12)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_13), "", Trim(g_rst_Genera!CAMPO_DET_13)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_14), "", Trim(g_rst_Genera!CAMPO_DET_14)); "|"; _
                  "[" & IIf(IsNull(g_rst_Genera!CAMPO_DET_15_1_1), "", Trim(g_rst_Genera!CAMPO_DET_15_1_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_1_2), "", Trim(g_rst_Genera!CAMPO_DET_15_1_2)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15_2_1), "", Trim(g_rst_Genera!CAMPO_DET_15_2_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_2_2), "", Trim(g_rst_Genera!CAMPO_DET_15_2_2)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15_3_1), "", Trim(g_rst_Genera!CAMPO_DET_15_3_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_3_2), "", Trim(g_rst_Genera!CAMPO_DET_15_3_2)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15_4_1), "", Trim(g_rst_Genera!CAMPO_DET_15_4_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_4_2), "", Trim(g_rst_Genera!CAMPO_DET_15_4_2)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15_5_1), "", Trim(g_rst_Genera!CAMPO_DET_15_5_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_5_2), "", Trim(g_rst_Genera!CAMPO_DET_15_5_2)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15_6_1), "", Trim(g_rst_Genera!CAMPO_DET_15_6_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_6_2), "", Trim(g_rst_Genera!CAMPO_DET_15_6_2)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15_7_1), "", Trim(g_rst_Genera!CAMPO_DET_15_7_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_7_2), "", Trim(g_rst_Genera!CAMPO_DET_15_7_2)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15_8_1), "", Trim(g_rst_Genera!CAMPO_DET_15_8_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_8_2), "", Trim(g_rst_Genera!CAMPO_DET_15_8_2)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15_9_1), "", Trim(g_rst_Genera!CAMPO_DET_15_9_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_9_2), "", Trim(g_rst_Genera!CAMPO_DET_15_9_2)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15_10_1), "", Trim(g_rst_Genera!CAMPO_DET_15_10_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_10_2), "", Trim(g_rst_Genera!CAMPO_DET_15_10_2)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15_11_1), "", Trim(g_rst_Genera!CAMPO_DET_15_11_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_15_11_2), "", Trim(g_rst_Genera!CAMPO_DET_15_11_2)); "]|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_37), "", Trim(g_rst_Genera!CAMPO_DET_37)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_38), "", Trim(g_rst_Genera!CAMPO_DET_38)); "|"; _
                  "[" & IIf(IsNull(g_rst_Genera!CAMPO_DET_39_1), "", Trim(g_rst_Genera!CAMPO_DET_39_1)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_39_2), "", Format(Trim(g_rst_Genera!CAMPO_DET_39_2), "0.00")); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_39_3), "", Trim(g_rst_Genera!CAMPO_DET_39_3)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_39_4), "", Format(Trim(g_rst_Genera!CAMPO_DET_39_4), "0.00")); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_39_5), "", Format(Trim(g_rst_Genera!CAMPO_DET_39_5), "0.00")); "]|"; vbCrLf;
                  
                  'IIf(IsNull(g_rst_Genera!CAMPO_DET_13), "", Trim(g_rst_Genera!CAMPO_DET_13)); "|";
            g_rst_Genera.MoveNext
         Loop
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         Print #r_int_NumRes, "##############################" & vbCrLf; 'vbCr
                     
      g_rst_Princi.MoveNext
      
      DoEvents
         
   Loop
   
               
'   If Not IsNull(g_rst_Princi!CAMPO_ADI_03) And g_rst_Princi!CAMPO_ADI_03 <> "" Then
'      Print #1, "ADI"; "|"; _
'             IIf(IsNull(g_rst_Princi!CAMPO_ADI_02), "", Trim(g_rst_Princi!CAMPO_ADI_02)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_ADI_03), "", Trim(g_rst_Princi!CAMPO_ADI_03)); "|"; _
'             IIf(IsNull(g_rst_Princi!CAMPO_ADI_04), "", Trim(g_rst_Princi!CAMPO_ADI_04)); "|";
'   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Close #1
   
   'Convertir a UTF8 sin BOM
   l_str_RutaArc = r_str_NomRes
   Call fs_Convertir_Utf8NoBom(r_str_NomRes)
   
   'Enviar el archivo en el ftp
   If fs_Cargar_Archivo(r_str_NomRes, r_str_Nombre) = True Then
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   Se envió correctamente el archivo :" & r_str_Nombre & ", procedimiento: fs_Generar_Archivo")
      
      'Leer el archivo enviado para determinar si se envió al Sftp
      Call fs_Leer_Archivo_EnvSFTP(r_str_NomRes)
      
   Else
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se cargó el archivo en SFTP, procedimiento: fs_Generar_Archivo")
   End If
   
   Exit Sub
   
MyError:

   Close #1
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Generar_Archivo")
   
End Sub
Private Sub fs_Generar_Archivo_07(ByVal p_DiaAct As String, ByVal p_TipDoc As String)

Dim r_str_NomRes     As String
Dim r_int_NumRes     As Integer
Dim r_str_Nombre     As String
Dim r_str_MtoLtr     As String
Dim r_str_Parame     As String

On Error GoTo MyError

   Screen.MousePointer = 11
   
   '*** GENERANDO REPORTE
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT DOCELE_CODIGO            , DOCELE_FECPRO            , DOCELE_FECAUT            , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_SERNUM          CAMPO_IDE_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_FECEMI          CAMPO_IDE_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_HOREMI          CAMPO_IDE_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_IDE_TIPMON          CAMPO_IDE_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_SERNUM          CAMPO_EMI_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_TIPDOC          CAMPO_EMI_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_NUMDOC          CAMPO_EMI_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_NOMCOM          CAMPO_EMI_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DENOMI          CAMPO_EMI_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_UBIGEO          CAMPO_EMI_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DIRCOM          CAMPO_EMI_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_URBANI          CAMPO_EMI_09      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_PROVIN          CAMPO_EMI_10      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DEPART          CAMPO_EMI_11      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_DISTRI          CAMPO_EMI_12      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_CODPAI          CAMPO_EMI_13      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_TELEMI          CAMPO_EMI_14      , "
   g_str_Parame = g_str_Parame & "        DOCELE_EMI_COREMI          CAMPO_EMI_15      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_SERNUM          CAMPO_REC_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_TIPDOC          CAMPO_REC_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_NUMDOC          CAMPO_REC_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DENOMI          CAMPO_REC_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DIRCOM          CAMPO_REC_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DEPART          CAMPO_REC_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_PROVIN          CAMPO_REC_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_DISTRI          CAMPO_REC_09      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_CODPAI          CAMPO_REC_10      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_TELREC          CAMPO_REC_11      , "
   g_str_Parame = g_str_Parame & "        DOCELE_REC_CORREC          CAMPO_REC_12      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_SERNUM          CAMPO_DRF_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_TIPDOC          CAMPO_DRF_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_NUMDOC          CAMPO_DRF_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_CODMOT          CAMPO_DRF_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_DRF_DESMOT          CAMPO_DRF_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_SERNUM          CAMPO_CAB_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEGRV   CAMPO_CAB_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEGRV   CAMPO_CAB_04      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEINA   CAMPO_CAB_05      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEINA   CAMPO_CAB_06      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_OPEEXO   CAMPO_CAB_07      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTVTA_OPEEXO   CAMPO_CAB_08      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIMP          CAMPO_CAB_18_1    , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_MTOIMP          CAMPO_CAB_18_2    , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_OTRCAR          CAMPO_CAB_19      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_CODIGO_TOTDSC   CAMPO_CAB_20      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTDSC          CAMPO_CAB_21      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_IMPTOT_DOCUME   CAMPO_CAB_22      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TOTANT          CAMPO_CAB_25      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_TIPOPE          CAMPO_CAB_26      , "
   g_str_Parame = g_str_Parame & "        DOCELE_CAB_LEYEND          CAMPO_CAB_27      , "
   g_str_Parame = g_str_Parame & "        DOCELE_ADI_SERNUM          CAMPO_ADI_02      , "
   g_str_Parame = g_str_Parame & "        DOCELE_ADI_TITADI          CAMPO_ADI_03      , "
   g_str_Parame = g_str_Parame & "        DOCELE_ADI_VALADI          CAMPO_ADI_04        "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_DOCELE "
   g_str_Parame = g_str_Parame & "  WHERE DOCELE_FLGENV = 0 "
   g_str_Parame = g_str_Parame & "    AND DOCELE_IDE_TIPDOC = '07' "
   g_str_Parame = g_str_Parame & "    AND DOCELE_DRF_NUMDOC IS NOT NULL "
   g_str_Parame = g_str_Parame & "  ORDER BY DOCELE_CODIGO ASC "

    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error en la consulta de Notas de Crédio en la tabla CNTBL_DOCELE, procedimiento: fs_Generar_Archivo")
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se encontró ningún registro de Notas de Crédito en la tabla CNTBL_DOCELE, procedimiento: fs_Generar_Archivo")
      Exit Sub
   End If
   
   'Creando Archivo
   r_str_Nombre = "20511904162-" & p_TipDoc & "-" & p_DiaAct & ".txt"
   r_str_NomRes = l_str_RutFacEnt & "20511904162-" & p_TipDoc & "-" & p_DiaAct & ".txt"
   r_int_NumRes = FreeFile
   
   Open r_str_NomRes For Output As r_int_NumRes
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
   
         Print #r_int_NumRes, "IDE"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_02), "", Trim(g_rst_Princi!CAMPO_IDE_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_03), "", Trim(g_rst_Princi!CAMPO_IDE_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_04), "", Trim(g_rst_Princi!CAMPO_IDE_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_IDE_06), "", Trim(g_rst_Princi!CAMPO_IDE_06)) & vbCrLf;
                   

         Print #r_int_NumRes, "EMI"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_02), "", Trim(g_rst_Princi!CAMPO_EMI_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_03), "", Trim(g_rst_Princi!CAMPO_EMI_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_04), "", Trim(g_rst_Princi!CAMPO_EMI_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_05), "", Trim(g_rst_Princi!CAMPO_EMI_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_06), "", Trim(g_rst_Princi!CAMPO_EMI_06)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_07), "", Trim(g_rst_Princi!CAMPO_EMI_07)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_08), "", Trim(g_rst_Princi!CAMPO_EMI_08)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_09), "", Trim(g_rst_Princi!CAMPO_EMI_09)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_10), "", Trim(g_rst_Princi!CAMPO_EMI_10)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_11), "", Trim(g_rst_Princi!CAMPO_EMI_11)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_12), "", Trim(g_rst_Princi!CAMPO_EMI_12)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_13), "", Trim(g_rst_Princi!CAMPO_EMI_13)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_14), "", Trim(g_rst_Princi!CAMPO_EMI_14)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_EMI_15), "", Trim(g_rst_Princi!CAMPO_EMI_15)) & vbCrLf;
                   
         
         Print #r_int_NumRes, "REC"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_02), "", Trim(g_rst_Princi!CAMPO_REC_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_03), "", Trim(g_rst_Princi!CAMPO_REC_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_04), "", Trim(g_rst_Princi!CAMPO_REC_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_05), "", Trim(g_rst_Princi!CAMPO_REC_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_06), "", Trim(g_rst_Princi!CAMPO_REC_06)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_07), "", Trim(g_rst_Princi!CAMPO_REC_07)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_08), "", Trim(g_rst_Princi!CAMPO_REC_08)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_09), "", Trim(g_rst_Princi!CAMPO_REC_09)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_10), "", Trim(g_rst_Princi!CAMPO_REC_10)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_11), "", Trim(g_rst_Princi!CAMPO_REC_11)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_REC_12), "", Trim(g_rst_Princi!CAMPO_REC_12)) & vbCrLf;
                   
                   
         If Not IsNull(g_rst_Princi!CAMPO_DRF_03) And g_rst_Princi!CAMPO_DRF_03 <> "" Then
            Print #r_int_NumRes, "DRF"; "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_02), "", Trim(g_rst_Princi!CAMPO_DRF_02)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_03), "", Trim(g_rst_Princi!CAMPO_DRF_03)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_04), "", Trim(g_rst_Princi!CAMPO_DRF_04)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_05), "", Trim(g_rst_Princi!CAMPO_DRF_05)); "|"; _
                   IIf(IsNull(g_rst_Princi!CAMPO_DRF_06), "", Trim(g_rst_Princi!CAMPO_DRF_06)) & vbCrLf;
                   
         End If
         
         r_str_MtoLtr = fs_NroEnLetras(g_rst_Princi!CAMPO_CAB_22)
         
         Print #r_int_NumRes, "CAB"; "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_02), "", Trim(g_rst_Princi!CAMPO_CAB_02)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_03), "", Trim(g_rst_Princi!CAMPO_CAB_03)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_04), "", Format(Trim(g_rst_Princi!CAMPO_CAB_04), "0.00")); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_05), "", Trim(g_rst_Princi!CAMPO_CAB_05)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_06), "", Trim(g_rst_Princi!CAMPO_CAB_06)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_07), "", Trim(g_rst_Princi!CAMPO_CAB_07)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_08), "", Format(Trim(g_rst_Princi!CAMPO_CAB_08), "0.00")); "|"; _
                  "[" & IIf(IsNull(g_rst_Princi!CAMPO_CAB_18_1), "", Trim(g_rst_Princi!CAMPO_CAB_18_1)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_18_2), "", Format(Trim(g_rst_Princi!CAMPO_CAB_18_2), "0.00")); "]|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_19), "", Trim(g_rst_Princi!CAMPO_CAB_19)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_20), "", Trim(g_rst_Princi!CAMPO_CAB_20)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_CAB_21), "", Trim(g_rst_Princi!CAMPO_CAB_21)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_22), "", Trim(g_rst_Princi!CAMPO_CAB_22)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_25), "", Trim(g_rst_Princi!CAMPO_CAB_25)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_26), "", Trim(g_rst_Princi!CAMPO_CAB_26)); "|"; _
                  IIf(IsNull(g_rst_Princi!CAMPO_CAB_27), "", Trim(g_rst_Princi!CAMPO_CAB_27) & "|" & r_str_MtoLtr & "]") & vbCrLf;
                  
                 
         'Generar detalle de las facturas
         moddat_g_str_CodGen = g_rst_Princi!DOCELE_CODIGO
            
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " SELECT DOCELEDET_CODIGO                              , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_SERNUM          CAMPO_DET_02    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_NUMITE          CAMPO_DET_03    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODPRD          CAMPO_DET_04    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DESPRD          CAMPO_DET_05    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CANTID          CAMPO_DET_06    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_UNIDAD          CAMPO_DET_07    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_VALUNI          CAMPO_DET_08    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PUNVTA          CAMPO_DET_09    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIMP          CAMPO_DET_10_1  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_MTOIMP          CAMPO_DET_10_2  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_TIPAFE          CAMPO_DET_10_3  , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_VALVTA          CAMPO_DET_11    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_VALREF          CAMPO_DET_12    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODSUN          CAMPO_DET_15    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODCON          CAMPO_DET_16    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_NROCON          CAMPO_DET_17    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_FECOTO   CAMPO_DET_18    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_FECOTO          CAMPO_DET_19    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_TIPPRE   CAMPO_DET_20    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_TIPPRE          CAMPO_DET_21    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_PARREG   CAMPO_DET_22    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PARREG          CAMPO_DET_23    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_PRIVIV   CAMPO_DET_24    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PRIVIV          CAMPO_DET_25    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODIGO_DIRCOM   CAMPO_DET_26    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DIRCOM          CAMPO_DET_27    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODUBI          CAMPO_DET_28    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_UBIGEO          CAMPO_DET_29    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODURB          CAMPO_DET_30    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_URBANI          CAMPO_DET_31    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODDPT          CAMPO_DET_32    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DEPART          CAMPO_DET_33    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODPRV          CAMPO_DET_34    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_PROVIN          CAMPO_DET_35    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_CODDIS          CAMPO_DET_36    , "
         g_str_Parame = g_str_Parame & "        DOCELEDET_DET_DISTRI          CAMPO_DET_37      "
         g_str_Parame = g_str_Parame & "   FROM CNTBL_DOCELEDET "
         g_str_Parame = g_str_Parame & "  WHERE DOCELEDET_CODIGO = " & CLng(moddat_g_str_CodGen)
         g_str_Parame = g_str_Parame & "  ORDER BY DOCELEDET_CODIGO, DOCELEDET_DET_NUMITE ASC "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se lee la consulta en la tabla CNTBL_DOCELEDET, Nro Ope:" & moddat_g_str_CodGen & ", procedimiento: fs_Generar_Archivo")
            Exit Sub
         End If
              
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            g_rst_Genera.Close
            Set g_rst_Genera = Nothing
            Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No hay ningun registro en la tabla CNTBL_DOCELEDET, Nro Ope:" & moddat_g_str_CodGen & ", procedimiento: fs_Generar_Archivo")
            Exit Sub
         End If
         
         g_rst_Genera.MoveFirst
         
         Do While Not g_rst_Genera.EOF
         
            Print #r_int_NumRes, "DET"; "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_02), "", Trim(g_rst_Genera!CAMPO_DET_02)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_03), "", Trim(g_rst_Genera!CAMPO_DET_03)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_04), "", Trim(g_rst_Genera!CAMPO_DET_04)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_05), "", Trim(g_rst_Genera!CAMPO_DET_05)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_06), "", Format(Trim(g_rst_Genera!CAMPO_DET_06), "0.000")); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_07), "", Trim(g_rst_Genera!CAMPO_DET_07)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_08), "", Trim(g_rst_Genera!CAMPO_DET_08)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_09), "", Trim(g_rst_Genera!CAMPO_DET_09)); "|"; _
                  "[" & IIf(IsNull(g_rst_Genera!CAMPO_DET_10_1), "", Trim(g_rst_Genera!CAMPO_DET_10_1)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_10_2), "", Format(Trim(g_rst_Genera!CAMPO_DET_10_2), "0.00")); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_10_3), "", Trim(g_rst_Genera!CAMPO_DET_10_3)); "]|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_11), "", Format(Trim(g_rst_Genera!CAMPO_DET_11), "0.00")); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_12), "", Trim(g_rst_Genera!CAMPO_DET_12)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_15), "", Trim(g_rst_Genera!CAMPO_DET_15)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_16), "", Trim(g_rst_Genera!CAMPO_DET_16)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_17), "", Trim(g_rst_Genera!CAMPO_DET_17)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_18), "", Trim(g_rst_Genera!CAMPO_DET_18)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_19), "", Trim(g_rst_Genera!CAMPO_DET_19)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_20), "", Trim(g_rst_Genera!CAMPO_DET_20)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_21), "", Trim(g_rst_Genera!CAMPO_DET_21)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_22), "", Trim(g_rst_Genera!CAMPO_DET_22)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_23), "", Trim(g_rst_Genera!CAMPO_DET_23)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_24), "", Trim(g_rst_Genera!CAMPO_DET_24)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_25), "", Trim(g_rst_Genera!CAMPO_DET_25)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_26), "", Trim(g_rst_Genera!CAMPO_DET_26)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_27), "", Trim(g_rst_Genera!CAMPO_DET_27)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_28), "", Trim(g_rst_Genera!CAMPO_DET_28)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_29), "", Trim(g_rst_Genera!CAMPO_DET_29)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_30), "", Trim(g_rst_Genera!CAMPO_DET_30)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_31), "", Trim(g_rst_Genera!CAMPO_DET_31)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_32), "", Trim(g_rst_Genera!CAMPO_DET_32)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_33), "", Trim(g_rst_Genera!CAMPO_DET_33)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_34), "", Trim(g_rst_Genera!CAMPO_DET_34)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_35), "", Trim(g_rst_Genera!CAMPO_DET_35)); "|"; IIf(IsNull(g_rst_Genera!CAMPO_DET_36), "", Trim(g_rst_Genera!CAMPO_DET_36)); "|"; _
                  IIf(IsNull(g_rst_Genera!CAMPO_DET_37), "", Trim(g_rst_Genera!CAMPO_DET_37)); vbCrLf;
                                    
            g_rst_Genera.MoveNext
         Loop
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         Print #r_int_NumRes, "##############################" & vbCrLf; 'vbCr
                     
      g_rst_Princi.MoveNext
      
      DoEvents
         
   Loop
               
'   If Not IsNull(g_rst_Princi!CAMPO_ADI_03) And g_rst_Princi!CAMPO_ADI_03 <> "" Then
'      Print #1, "ADI"; "|"; _
'             IIf(IsNull(g_rst_Princi!CAMPO_ADI_02), "", Trim(g_rst_Princi!CAMPO_ADI_02)); "|"; IIf(IsNull(g_rst_Princi!CAMPO_ADI_03), "", Trim(g_rst_Princi!CAMPO_ADI_03)); "|"; _
'             IIf(IsNull(g_rst_Princi!CAMPO_ADI_04), "", Trim(g_rst_Princi!CAMPO_ADI_04)); "|";
'   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Close #1
   
   'Convertir en UTF8 NO BOM
   l_str_RutaArc = r_str_NomRes
   Call fs_Convertir_Utf8NoBom(r_str_NomRes)
   
   'Enviar el archivo en el ftp
   If fs_Cargar_Archivo(r_str_NomRes, r_str_Nombre) = True Then
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   Se envió correctamente el archivo :" & r_str_Nombre & ", procedimiento: fs_Generar_Archivo")
      
      'Leer el archivo enviado para determinar si se envió al Sftp
      Call fs_Leer_Archivo_EnvSFTP(r_str_NomRes)
      
   Else
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se cargó el archivo en SFTP, procedimiento: fs_Generar_Archivo")
   End If
   
   Exit Sub
   
MyError:

   Close #1
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Generar_Archivo")
   
End Sub
Private Sub fs_Leer_Archivo_EnvSFTP(p_sFile)
Dim r_str_Cadena     As String
Dim r_arr_NumFac()   As moddat_tpo_Genera
Dim r_lng_Contad     As Long
Dim r_str_Parame     As String
Dim r_str_CadAux     As String

On Error GoTo Err

   ReDim r_arr_NumFac(0)
   Open p_sFile For Input As #1


   Do While Not EOF(1)
      Line Input #1, r_str_Cadena
      
      If InStr(r_str_Cadena, "IDE") > 0 Then

         ReDim Preserve r_arr_NumFac(UBound(r_arr_NumFac) + 1)
      
         r_str_CadAux = Trim(Mid(r_str_Cadena, InStr(r_str_Cadena, "IDE|") + Len("IDE|")))
         r_str_CadAux = Trim(Mid(r_str_CadAux, 1, InStr(r_str_CadAux, "|") - 1))
         
         r_arr_NumFac(UBound(r_arr_NumFac)).Genera_Codigo = r_str_CadAux
      End If
   Loop
Close #1

   'Actualiza campo DOCELE_FLGENV para identificar cuales se enviaron en el archivo
   For r_lng_Contad = 0 To UBound(r_arr_NumFac)
      If Len(Trim(r_arr_NumFac(r_lng_Contad).Genera_Codigo)) > 0 Then
         r_str_Parame = ""
         r_str_Parame = r_str_Parame & "UPDATE CNTBL_DOCELE SET DOCELE_FLGENV = 1 "
         r_str_Parame = r_str_Parame & " WHERE DOCELE_IDE_SERNUM = '" & r_arr_NumFac(r_lng_Contad).Genera_Codigo & "' "
         
         If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar DOCELE_FLGENV de la tabla CNTBL_DOCELE, procedimiento: fs_Leer_Archivo_EnvSFTP")
            Exit Sub
         End If
      End If
   Next
   Exit Sub
   
Err:
Close #1
Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Leer_Archivo_EnvSFTP")
Err.Clear

End Sub
Private Sub fs_Convertir_Utf8NoBom(p_sFile)

Dim UTFStream              As New ADODB.Stream
Dim ANSIStream             As New ADODB.Stream
Dim BinaryStream           As New ADODB.Stream

On Error GoTo MyError
    ANSIStream.Type = adTypeText
    ANSIStream.Mode = adModeReadWrite
    ANSIStream.Charset = "iso-8859-1"
    ANSIStream.Open
    ANSIStream.LoadFromFile p_sFile   'ANSI File
    
    UTFStream.Type = adTypeText
    UTFStream.Mode = adModeReadWrite
    UTFStream.Charset = "UTF-8"
    UTFStream.Open
    ANSIStream.CopyTo UTFStream
    

    UTFStream.Position = 3 'skip BOM
    BinaryStream.Type = adTypeBinary
    BinaryStream.Mode = adModeReadWrite
    BinaryStream.Open

    'Strips BOM (first 3 bytes)
    UTFStream.CopyTo BinaryStream

    BinaryStream.SaveToFile p_sFile, adSaveCreateOverWrite
    BinaryStream.Flush
    BinaryStream.Close
    Exit Sub
MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", Convertir archivo a UTF8, procedimiento: fs_Genera_Utf8NoBom")
End Sub

Private Sub cmd_ConFac_Click()
   frm_RptSun_08.Show 1
End Sub

Private Sub cmd_Proces_Click()

Dim r_str_RutLog     As String

   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   MsgBox "Esta opción descarga los archivos previamente cargados", vbExclamation, modgen_g_str_NomPlt
   
   If MsgBox("¿Está seguro de descargar los archivos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Crear Archivo LOG del Proceso
   l_str_NomLOG = UCase(App.EXEName) & "_D_" & Format(date, "yyyymmdd") & ".LOG"
   l_int_NumLOG = FreeFile
   
   r_str_RutLog = Replace(moddat_g_str_RutFac, "\Fact", "\Logs")
   
   If gf_Existe_Archivo(r_str_RutLog & "\", l_str_NomLOG) Then
      Kill r_str_RutLog & "\" & l_str_NomLOG
      DoEvents
   End If
   
   l_str_RutaLg = r_str_RutLog & "\" & l_str_NomLOG
   
   'moddat_g_str_RutLoc
   l_str_RutFacEnt = moddat_g_str_RutFac & "\entrada\"
   l_str_RutFacRep = moddat_g_str_RutFac & "\reportes\"
   l_str_RutFacAce = moddat_g_str_RutFac & "\reportes\aceptados\"
   l_str_RutFacRec = moddat_g_str_RutFac & "\reportes\rechazados\"

   'Crear la Carpeta Reportes
   Set l_fsobj = New FileSystemObject
   If l_fsobj.FolderExists(l_str_RutFacRep) = False Then
      l_fsobj.CreateFolder (l_str_RutFacRep)
   End If
   
   'Crear la Carpeta Reportes
   Set l_fsobj = New FileSystemObject
   If l_fsobj.FolderExists(l_str_RutFacAce) = False Then
      l_fsobj.CreateFolder (l_str_RutFacAce)
   End If
   
   'Crear la Carpeta Reportes
   Set l_fsobj = New FileSystemObject
   If l_fsobj.FolderExists(l_str_RutFacRec) = False Then
      l_fsobj.CreateFolder (l_str_RutFacRec)
   End If
   
   Open l_str_RutaLg For Output As l_int_NumLOG
   Close #l_int_NumLOG
   
   Call fs_Escribir_Linea(l_str_RutaLg, "")
   Call fs_Escribir_Linea(l_str_RutaLg, "Proceso           : " & modgen_g_str_NomPlt)
   Call fs_Escribir_Linea(l_str_RutaLg, "Proceso           : " & modgen_g_str_NomPlt)
   Call fs_Escribir_Linea(l_str_RutaLg, "Nombre Ejecutable : " & UCase(App.EXEName))
   Call fs_Escribir_Linea(l_str_RutaLg, "Número Revisión   : " & modgen_g_str_NumRev)
   Call fs_Escribir_Linea(l_str_RutaLg, "Nombre PC         : " & modgen_g_str_NombPC)
   Call fs_Escribir_Linea(l_str_RutaLg, "Origen Datos      : " & moddat_g_str_NomEsq & " - " & moddat_g_str_EntDat)
   Call fs_Escribir_Linea(l_str_RutaLg, "")
   Call fs_Escribir_Linea(l_str_RutaLg, "Inicio Proceso    : " & Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss"))
   Call fs_Escribir_Linea(l_str_RutaLg, "")
      
   'Descargando los archivos del SFTP
   If fs_Descargar_Archivo(l_str_FecCar) = True Then
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   Se descargaron correctamente los archivos, procedimiento: fs_Descargar_Archivo")
      
      'Leyendo las respuestas descargadas del SFTP
      Call fs_Leer_RptArc
     
   End If
   
   'Cerrando Archivo LOG del Proceso
   Call fs_Escribir_Linea(l_str_RutaLg, "")
   Call fs_Escribir_Linea(l_str_RutaLg, "Fecha Proceso     : " & Format(date, "dd/mm/yyyy"))
   Call fs_Escribir_Linea(l_str_RutaLg, "Fin Proceso       : " & Format(date, "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss"))
   Call fs_Escribir_Linea(l_str_RutaLg, "")
   
   'Para enviar Correo Electrónico
   'Call fs_Envia_CorEle_User
   'Call fs_Envia_CorEle_LOG
   
   MsgBox "Proceso finalizado.", vbInformation, modgen_g_str_NomPlt
    
   Screen.MousePointer = 0
End Sub
Private Sub fs_Leer_RptArc()

Dim r_arr_RptFac()         As moddat_tpo_Genera
Dim r_lng_Contad           As Long
Dim r_str_Parame           As String
Dim r_str_CadAux           As String
Dim r_str_arcdir           As String
Dim r_str_NumFac           As String

Dim xml_document           As DOMDocument
Dim oNode                  As IXMLDOMNode
Dim oSNode                 As IXMLDOMNode
Dim oSsNode                As IXMLDOMNode
'Dim oAttr                  As IXMLDOMAttribute

Dim cbc_ResponseCode       As String
Dim cbc_Description        As String
Dim cbc_Status             As String
Dim cac_DocumentReferense  As String
Dim cac_IssuerParty        As String
Dim cac_RecipientParty     As String

Dim r_shf_SHFileOp         As SHFILEOPSTRUCT
Dim r_str_FicOri           As String
Dim r_str_FicDes           As String

On Error GoTo Err
   
   Screen.MousePointer = 11
   
   ReDim r_arr_RptFac(0)
   
   r_str_arcdir = Dir(l_str_RutFacRep & "*.xml")
   
   While r_str_arcdir <> ""
      
      If InStr(r_str_arcdir, "R-") = 0 Then
         r_str_arcdir = Dir
      Else
      
         Set xml_document = New DOMDocument
         xml_document.Load l_str_RutFacRep & r_str_arcdir
            
         If xml_document.documentElement Is Nothing Then
            Exit Sub
         End If

         For Each oNode In xml_document.documentElement.childNodes
                   
            If oNode.nodeName = "cac:DocumentResponse" Then
                For Each oSNode In oNode.childNodes
                  If oSNode.nodeName = "cac:Response" Then
                     For Each oSsNode In oSNode.childNodes
                        If oSsNode.nodeName = "cbc:ResponseCode" Then
                           cbc_ResponseCode = oSsNode.Text
                        ElseIf oSsNode.nodeName = "cbc:Description" Then
                           cbc_Description = oSsNode.Text
                        ElseIf oSsNode.nodeName = "cac:Status" Then
                           cbc_Status = cbc_Status & " " & oSsNode.Text & vbCrLf
                        End If
                     Next oSsNode
                  ElseIf oSNode.nodeName = "cac:DocumentReference" Then
                     cac_DocumentReferense = oSNode.Text
                  ElseIf oSNode.nodeName = "cac:IssuerParty" Then
                     cac_IssuerParty = oSNode.Text
                  ElseIf oSNode.nodeName = "cac:RecipientParty" Then
                     cac_RecipientParty = oSNode.Text
                  End If
                Next oSNode

            End If
            
         Next oNode
      
         ReDim Preserve r_arr_RptFac(UBound(r_arr_RptFac) + 1)
         r_arr_RptFac(UBound(r_arr_RptFac)).Genera_Codigo = r_str_arcdir
         If InStr(r_str_arcdir, "F00") > 0 Then
            r_arr_RptFac(UBound(r_arr_RptFac)).Genera_ConNDo = Trim(Mid(r_str_arcdir, InStr(r_str_arcdir, "F00")))
         ElseIf InStr(r_str_arcdir, "B00") > 0 Then
            r_arr_RptFac(UBound(r_arr_RptFac)).Genera_ConNDo = Trim(Mid(r_str_arcdir, InStr(r_str_arcdir, "B00")))
         End If

         If InStr(cbc_Description, "ha sido aceptado") = 0 Then 'cbc_ResponseCode <> "" And cbc_Description <> "" And cbc_Status = "" And
            r_arr_RptFac(UBound(r_arr_RptFac)).Genera_FlgAso = 0
         Else
            r_arr_RptFac(UBound(r_arr_RptFac)).Genera_FlgAso = 1
         End If
         r_arr_RptFac(UBound(r_arr_RptFac)).Genera_Refere = cbc_Description & vbCrLf & cbc_Status
         
         r_str_arcdir = Dir
         cbc_ResponseCode = Empty
         cbc_Description = Empty
         cbc_Status = Empty
         cac_DocumentReferense = Empty
         cac_IssuerParty = Empty
         cac_RecipientParty = Empty
      
     End If
     DoEvents:
   Wend
   
   'Actualiza el campo DOCELE_FLGRPT si ha sido aceptada la factura
   For r_lng_Contad = 0 To UBound(r_arr_RptFac)
   
      If Len(Trim(r_arr_RptFac(r_lng_Contad).Genera_Codigo)) > 0 Then
         
         r_str_NumFac = Replace(r_arr_RptFac(r_lng_Contad).Genera_ConNDo, "QA", "00")
         r_str_NumFac = Replace(r_str_NumFac, ".xml", "")
            
         r_str_Parame = ""
         r_str_Parame = r_str_Parame & " UPDATE CNTBL_DOCELE SET "
         r_str_Parame = r_str_Parame & "        DOCELE_FECAUT = " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "        , "
         r_str_Parame = r_str_Parame & "        DOCELE_FLGRPT = " & CInt(Trim(r_arr_RptFac(r_lng_Contad).Genera_FlgAso)) & "  , "
         r_str_Parame = r_str_Parame & "        DOCELE_OBSERV = '" & Trim(r_arr_RptFac(r_lng_Contad).Genera_Refere) & "'      , "
         r_str_Parame = r_str_Parame & "        SEGUSUACT = '" & modgen_g_str_CodUsu & "'                                     , "
         r_str_Parame = r_str_Parame & "        SEGFECACT = " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & "            , "
         r_str_Parame = r_str_Parame & "        SEGHORACT = " & Format(Time, "HHMMSS") & "                                    , "
         r_str_Parame = r_str_Parame & "        SEGPLTACT = '" & UCase(App.EXEName) & "'                                      , "
         r_str_Parame = r_str_Parame & "        SEGTERACT = '" & modgen_g_str_NombPC & "'                                     , "
         r_str_Parame = r_str_Parame & "        SEGSUCACT = '" & modgen_g_str_CodSuc & "'                                       "
         r_str_Parame = r_str_Parame & " WHERE DOCELE_IDE_SERNUM = '" & r_str_NumFac & "' "

         If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 2) Then
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Error no se puede actualizar DOCELE_FLGENV de la tabla CNTBL_DOCELE, procedimiento: fs_Leer_RptArc")
            Exit Sub
         End If
         
         
         DoEvents:
         
         ' Mover los archivos de respuesta en la carpeta aceptados o rechazados según sea el caso
         If Trim(r_arr_RptFac(r_lng_Contad).Genera_FlgAso) = 1 Then
            r_str_FicOri = l_str_RutFacRep & r_arr_RptFac(r_lng_Contad).Genera_Codigo & vbNullChar & vbNullChar
            r_str_FicDes = l_str_RutFacAce & r_arr_RptFac(r_lng_Contad).Genera_Codigo & vbNullChar & vbNullChar
         Else
            r_str_FicOri = l_str_RutFacRep & r_arr_RptFac(r_lng_Contad).Genera_Codigo & vbNullChar & vbNullChar
            r_str_FicDes = l_str_RutFacRec & r_arr_RptFac(r_lng_Contad).Genera_Codigo & vbNullChar & vbNullChar
         End If

         With r_shf_SHFileOp
             .wFunc = FO_MOVE
             .fFlags = 8                                       'iFlags Cambiar el nombre, si el destino ya existe (FOF_RENAMEONCOLLISION)
             .hwnd = Me.hwnd
             .pFrom = r_str_FicOri
             .pTo = r_str_FicDes
             '.lpszProgressTitle = "Moviendo los ficheros especificados"
         End With
      
         Call SHFileOperation(r_shf_SHFileOp)
         
         
         'Mover los archivos pdf
         'r_str_CadAux = "20511904162-01-" & r_arr_RptFac(r_lng_Contad).Genera_ConNDo
         r_str_CadAux = Replace(r_arr_RptFac(r_lng_Contad).Genera_Codigo, "R-", "")
         If Trim(r_arr_RptFac(r_lng_Contad).Genera_FlgAso) = 1 Then
            r_str_FicOri = l_str_RutFacRep & r_str_CadAux & vbNullChar & vbNullChar
            r_str_FicDes = l_str_RutFacAce & r_str_CadAux & vbNullChar & vbNullChar
         Else
            r_str_FicOri = l_str_RutFacRep & r_str_CadAux & vbNullChar & vbNullChar
            r_str_FicDes = l_str_RutFacRec & r_str_CadAux & vbNullChar & vbNullChar
         End If
         
         With r_shf_SHFileOp
             .wFunc = FO_MOVE
             .fFlags = 8 '0                                    'iFlags Cambiar el nombre, si el destino ya existe (FOF_RENAMEONCOLLISION)
             .hwnd = Me.hwnd
             .pFrom = r_str_FicOri
             .pTo = r_str_FicDes
             '.lpszProgressTitle = "Moviendo los ficheros especificados"
         End With
      
         Call SHFileOperation(r_shf_SHFileOp)
         
         'Mover los archivos xml
         r_str_CadAux = Replace(r_str_CadAux, "xml", "pdf")
         If Trim(r_arr_RptFac(r_lng_Contad).Genera_FlgAso) = 1 Then
            r_str_FicOri = l_str_RutFacRep & r_str_CadAux & vbNullChar & vbNullChar
            r_str_FicDes = l_str_RutFacAce & r_str_CadAux & vbNullChar & vbNullChar
         Else
            r_str_FicOri = l_str_RutFacRep & r_str_CadAux & vbNullChar & vbNullChar
            r_str_FicDes = l_str_RutFacRec & r_str_CadAux & vbNullChar & vbNullChar
         End If
         
         With r_shf_SHFileOp
             .wFunc = FO_MOVE
             .fFlags = 8 '0                                      'iFlags Cambiar el nombre, si el destino ya existe (FOF_RENAMEONCOLLISION)
             .hwnd = Me.hwnd
             .pFrom = r_str_FicOri
             .pTo = r_str_FicDes
             '.lpszProgressTitle = "Moviendo los ficheros especificados"
         End With
      
         Call SHFileOperation(r_shf_SHFileOp)
         
      End If
   Next r_lng_Contad
      
   Exit Sub
   
Err:

Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Leer_RptArc")
Err.Clear

End Sub
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
Private Function fs_Descargar_Archivo(p_DiaAct) As Boolean

Dim r_str_ResPta     As Long
Dim r_dbl_Puerto     As Long
Dim r_str_Server     As String
Dim r_str_handle     As String
Dim r_str_NomRes     As String
Dim r_lng_Contad     As Long
Dim r_int_ExiArc     As Integer
Dim r_str_DirRem     As String
Dim r_lng_mode       As Long
Dim r_lng_recurse    As Long
Dim r_key            As New ChilkatSshKey
Dim r_privKey        As String
Dim r_str_success    As String

   On Error GoTo MyError:

   fs_Descargar_Archivo = False
   Screen.MousePointer = 11
   
   Set r_chi_sftp = New ChilkatSFtp
   
   Call fs_Leer_FacDescargar
   
   r_str_ResPta = r_chi_sftp.UnlockComponent("30")

   If (r_str_ResPta <> 1) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & vbCrLf & " , procedimiento: fs_Descargar_Archivo")
      Exit Function
   End If

   r_chi_sftp.ConnectTimeoutMs = 5000
   r_chi_sftp.IdleTimeoutMs = 10000

   '  Producción:
   '  Sftp.escondatagate.net (puerto 6022)
   '  Calidad:
   '  Sftpqa.escondatagate.net (puerto 3022)
   
   r_str_Server = "Sftp.escondatagate.net"
   r_dbl_Puerto = 6022

'   r_str_Server = "Sftpqa.escondatagate.net"
'   r_dbl_Puerto = 3022

   r_str_ResPta = r_chi_sftp.Connect(r_str_Server, r_dbl_Puerto)
   If (r_str_ResPta <> 1) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & vbCrLf & " , procedimiento: fs_Descargar_Archivo")
      Exit Function
   End If

   'clave pública
   r_privKey = r_key.LoadText(moddat_g_str_RutFac & "\" & "id_rsa.ppk")
   
   If (r_key.LastMethodSuccess <> 1) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_key.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
      Exit Function
   End If
   
   r_str_success = r_key.FromOpenSshPrivateKey(r_privKey)
   If (r_str_success <> 1) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_key.LastErrorText & " , procedimiento: fs_Cargar_Archivo")
      Exit Function
   End If
   '---

   r_str_ResPta = r_chi_sftp.AuthenticatePw("micasi02", "Micasi2018*")
   If (r_str_ResPta <> 1) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & vbCrLf & " , procedimiento: fs_Descargar_Archivo")
      Exit Function
   End If

   r_str_ResPta = r_chi_sftp.InitializeSftp()
   If (r_str_ResPta <> 1) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & vbCrLf & " , procedimiento: fs_Descargar_Archivo")
      Exit Function
   End If

   'Mode = 0: Hace que SyncTreeDownload descargue todos los archivos.
   'recursive: No desciende recursivamente por el árbol del directorio remoto. Simplemente descarga todos los archivos en el directorio especificado.
'   r_str_DirRem = "/WWW/reportes/"
'   r_lng_mode = 0
'   r_lng_recurse = 0

'   r_str_ResPta = r_chi_sftp.SyncTreeDownload(r_str_DirRem, l_str_RutFacRep, r_lng_mode, r_lng_recurse)
'
'   If (r_str_ResPta <> 1) Then
'       Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & vbCrLf & " , procedimiento: fs_Descargar_Archivo")
'   End If
   
   For r_lng_Contad = 0 To UBound(l_arr_NumFac)

      If Len(Trim(l_arr_NumFac(r_lng_Contad).Genera_Codigo)) > 0 Then
   
         'Verifica si el archivo existe en el servidor remoto
         r_int_ExiArc = r_chi_sftp.FileExists("/WWW/reportes/" & l_arr_NumFac(r_lng_Contad).Genera_Codigo, 0)
   
         If r_int_ExiArc = 1 Then
   
            r_str_handle = r_chi_sftp.OpenFile("/WWW/reportes/" & l_arr_NumFac(r_lng_Contad).Genera_Codigo, "readOnly", "openExisting")
   
            If (r_chi_sftp.LastMethodSuccess <> 1) Then
               Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & vbCrLf & " , procedimiento: fs_Descargar_Archivo")
            Else
               r_str_ResPta = r_chi_sftp.DownloadFile(r_str_handle, l_str_RutFacRep & l_arr_NumFac(r_lng_Contad).Genera_Codigo)   'moddat_g_str_RutFac
   
               If (r_str_ResPta <> 1) Then
                  Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & vbCrLf & " , procedimiento: fs_Descargar_Archivo")
               End If
   
               r_str_ResPta = r_chi_sftp.CloseHandle(r_str_handle)
               If (r_str_ResPta <> 1) Then
                  Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & " " & r_chi_sftp.LastErrorText & vbCrLf & " , procedimiento: fs_Descargar_Archivo")
               End If
            End If
         Else
            Call fs_Escribir_Linea(l_str_RutaLg, "ERR  No existe el archivo: " & l_arr_NumFac(r_lng_Contad).Genera_Codigo & ", procedimiento: fs_Descargar_Archivo")
         End If
      End If
      DoEvents
   Next
   
   fs_Descargar_Archivo = True
   Exit Function

MyError:
   Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Descargar_Archivo")

End Function
Private Sub fs_Leer_FacDescargar()
Dim r_str_CadAux     As String
   
   On Error GoTo MyError
   
   ReDim l_arr_NumFac(0)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "     SELECT DOCELE_IDE_SERNUM  AS NUMERO_FACTURA,  DOCELE_IDE_TIPDOC AS TIPO_DOCUM "
   g_str_Parame = g_str_Parame & "       FROM CNTBL_DOCELE  "
   g_str_Parame = g_str_Parame & "      WHERE DOCELE_FLGENV = 1 "
   g_str_Parame = g_str_Parame & "        AND DOCELE_FLGRPT = 0 "
   
   If cmb_TipDoc.ListIndex <> 0 Then
      g_str_Parame = g_str_Parame & "        AND DOCELE_IDE_TIPDOC = '" & Format(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex), "00") & "'"
   End If
   
   g_str_Parame = g_str_Parame & "      ORDER BY DOCELE_CODIGO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call fs_Escribir_Linea(l_str_RutaLg, "ERR   No se pudo leer la tabla de destinatarios de correo, procedimiento: fs_Leer_FacDescargar")
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Call fs_Escribir_Linea(l_str_RutaLg, "VAL   No se han encontrado facturas para ser descargadas, procedimiento: fs_Leer_FacDescargar")
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      'Archivos de respuesta
      ReDim Preserve l_arr_NumFac(UBound(l_arr_NumFac) + 1)
      r_str_CadAux = g_rst_Princi!NUMERO_FACTURA
      r_str_CadAux = "R-20511904162-" & g_rst_Princi!TIPO_DOCUM & "-" & r_str_CadAux & ".xml"
      l_arr_NumFac(UBound(l_arr_NumFac)).Genera_Codigo = r_str_CadAux
         
      'Archivos pdf
      ReDim Preserve l_arr_NumFac(UBound(l_arr_NumFac) + 1)
      r_str_CadAux = g_rst_Princi!NUMERO_FACTURA
      r_str_CadAux = "20511904162-" & g_rst_Princi!TIPO_DOCUM & "-" & r_str_CadAux & ".xml"
      l_arr_NumFac(UBound(l_arr_NumFac)).Genera_Codigo = r_str_CadAux
         
      'Archivos xml
      ReDim Preserve l_arr_NumFac(UBound(l_arr_NumFac) + 1)
      r_str_CadAux = g_rst_Princi!NUMERO_FACTURA
      r_str_CadAux = "20511904162-" & g_rst_Princi!TIPO_DOCUM & "-" & r_str_CadAux & ".pdf"
      l_arr_NumFac(UBound(l_arr_NumFac)).Genera_Codigo = r_str_CadAux
      
      g_rst_Princi.MoveNext
   Loop
   Exit Sub
   
MyError:
Call fs_Escribir_Linea(l_str_RutaLg, "ERR   Nro: " & Err.Number & " " & Err.Description & ", procedimiento: fs_Leer_FacDescargar")

End Sub
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt

   Call fs_Limpia
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Limpia()
   ipp_FecIni.Text = Format(Now, "dd/mm/yyyy")
End Sub

