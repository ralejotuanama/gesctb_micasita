VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frm_Ctb_PagCom_08 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10470
   Icon            =   "GesCtb_frm_228.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   4290
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10545
      _Version        =   65536
      _ExtentX        =   18600
      _ExtentY        =   7567
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
         TabIndex        =   8
         Top             =   60
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
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
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   9690
            Top             =   30
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   495
            Left            =   630
            TabIndex        =   9
            Top             =   60
            Width           =   5445
            _Version        =   65536
            _ExtentX        =   9604
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Imprimir Chueque"
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
            Picture         =   "GesCtb_frm_228.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   60
         TabIndex        =   10
         Top             =   780
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
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
         Begin VB.CommandButton cmd_Config 
            Height          =   585
            Left            =   630
            Picture         =   "GesCtb_frm_228.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Impresora Predeterminada"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9750
            Picture         =   "GesCtb_frm_228.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Print 
            Height          =   585
            Left            =   30
            Picture         =   "GesCtb_frm_228.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Imprimir Cheque"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2685
         Left            =   60
         TabIndex        =   11
         Top             =   1500
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
         _ExtentY        =   4736
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
         Begin VB.TextBox txt_CodOrigen 
            Height          =   315
            Left            =   8130
            MaxLength       =   500
            TabIndex        =   20
            Top             =   690
            Width           =   1605
         End
         Begin VB.TextBox txt_NumChq 
            Height          =   315
            Left            =   1410
            MaxLength       =   13
            TabIndex        =   0
            Top             =   380
            Width           =   1605
         End
         Begin VB.TextBox txt_NomDe 
            Height          =   315
            Left            =   1410
            MaxLength       =   100
            TabIndex        =   2
            Top             =   1065
            Width           =   8325
         End
         Begin EditLib.fpDateTime ipp_FecChq 
            Height          =   315
            Left            =   1410
            TabIndex        =   1
            Top             =   720
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2831
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
         Begin Threed.SSPanel pnl_Import 
            Height          =   315
            Left            =   1410
            TabIndex        =   3
            Top             =   1750
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin Threed.SSPanel pnl_ImpLet 
            Height          =   315
            Left            =   1410
            TabIndex        =   4
            Top             =   2100
            Width           =   8325
            _Version        =   65536
            _ExtentX        =   14684
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   1410
            TabIndex        =   18
            Top             =   1410
            Width           =   8325
            _Version        =   65536
            _ExtentX        =   14684
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   1485
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Importe Letras:"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   2190
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nro Cheque:"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   465
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Importe:"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   1830
            Width           =   570
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "A Nombre de:"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   795
            Width           =   495
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   90
            Width           =   510
         End
      End
   End
End
Attribute VB_Name = "frm_Ctb_PagCom_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_CodOrigen As String
Dim l_NomOrigen As String

Private Sub cmd_Config_Click()
   CommonDialog1.Flags = &H40&
   CommonDialog1.ShowPrinter
End Sub

Private Sub cmd_Print_Click()
'Dim Prt As Printer

   If Trim(txt_NumChq.Text) = "" Then
       MsgBox "Tiene que digitar el número de cheque.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(txt_NumChq)
       Exit Sub
   End If
   If Trim(txt_NomDe.Text) = "" Then
       MsgBox "Tiene que digitar a nombre de quien va el cheque.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(txt_NomDe)
       Exit Sub
   End If
'   For Each Prt In Printers
'       If InStr(UCase(Prt.DeviceName), "CHEQUE") > 0 Then
'          Set Printer = Prt
'          Exit For
'       End If
'   Next
    
   If InStr(UCase(Printer.DeviceName), "CHEQUE") = 0 Then
      MsgBox "Debe de seleccionar el dispositivo de impresora de cheques.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   If MsgBox("¿Esta seguro de imprimir el cheque?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
'      For Each Prt In Printers
'          If InStr(UCase(Prt.DeviceName), "CONTABILIDAD") > 0 Then
'             Set Printer = Prt
'             Exit For
'          End If
'      Next
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenTxt_Chq
   Screen.MousePointer = 0
   
'   For Each Prt In Printers
'       If InStr(UCase(Prt.DeviceName), "CONTABILIDAD") > 0 Then
'          Set Printer = Prt
'          Exit For
'       End If
'   Next
End Sub

Public Sub fs_GenTxt_Chq()
Dim r_str_NomPrv    As String
Dim r_str_ImpSol    As String
Dim r_str_CadAux    As String
Dim r_str_CadAux2   As String
Dim r_str_Parame    As String
Dim r_rst_Genera    As ADODB.Recordset

   r_str_NomPrv = ""
   r_str_ImpSol = ""
   
   Printer.Font = "ARIAL"
   Printer.FontSize = 11
   Printer.FontBold = True

   r_str_NomPrv = Trim(txt_NomDe.Text)
   r_str_ImpSol = Format(CDbl(pnl_Import.Caption), "###,###,##0.00")
   r_str_CadAux2 = Format(CDbl(pnl_Import.Caption), "###,###,##0.00")
   
   Printer.FontSize = 10
   Printer.Print Tab(5); ""
   Printer.FontSize = 1
   Printer.Print Tab(5); ""
   Printer.FontSize = 11
   Printer.Print Tab(37); " " & Format(ipp_FecChq.Text, "dd     mm    yyyy") & "          " & r_str_ImpSol
   Printer.FontSize = 9
   Printer.Print Tab(5); ""
   Printer.FontSize = 6
   Printer.Print Tab(5); ""
   Printer.FontSize = 9
   Printer.Print Tab(5); ""
   Printer.Print Tab(15); r_str_NomPrv
   r_str_CadAux = fs_NumLetra(Left(r_str_CadAux2, Len(r_str_CadAux2) - 3))
   r_str_CadAux = r_str_CadAux & " CON " & Right(Format(CDbl(pnl_Import.Caption), "###,###,##0.00"), 2) & "/100"
   Printer.FontSize = 8
   Printer.Print Tab(5); ""
   Printer.FontSize = 9
   Printer.Print Tab(2); r_str_CadAux
      
   Printer.EndDoc
  
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "USP_CNTBL_TABLA_LOG ( "
   r_str_Parame = r_str_Parame & "1, "                                       'TABLOG_TIPPRO(1= CHEQUE)
   r_str_Parame = r_str_Parame & "'" & Trim(txt_CodOrigen.Tag) & "', "        'TABLOG_CODIGO(CODIGO_ORIGEN)
   r_str_Parame = r_str_Parame & "'" & Trim(txt_CodOrigen.Text) & "', "      'TABLOG_NOMBRE(NOMBRE_MODULO)
   r_str_Parame = r_str_Parame & Format(ipp_FecChq.Text, "yyyymmdd") & ", "  'TABLOG_FECHA(FECHA CHEQUE)
   r_str_Parame = r_str_Parame & "'" & Trim(txt_NumChq.Text) & "', "         'TABLOG_VALCAD01(NUMERO CHEQUE)
   r_str_Parame = r_str_Parame & "'" & Trim(txt_NomDe.Text) & "', "          'TABLOG_VALCAD02(A NOMBRE DE)
   r_str_Parame = r_str_Parame & "'" & Trim(pnl_Moneda.Caption) & "', "      'TABLOG_VALCAD03(MONEDA)
   r_str_Parame = r_str_Parame & "'" & Trim(pnl_ImpLet.Caption) & "', "      'TABLOG_VALCAD04(NUMERO A LETRAS)
   r_str_Parame = r_str_Parame & "'', "  'TABLOG_VALCAD05
   r_str_Parame = r_str_Parame & "'', "  'TABLOG_VALCAD06
   r_str_Parame = r_str_Parame & "'', "  'TABLOG_VALCAD07
   r_str_Parame = r_str_Parame & "'', "  'TABLOG_VALCAD08
   r_str_Parame = r_str_Parame & CDbl(pnl_Import.Caption) & ", " 'TABLOG_VALNUM01
   r_str_Parame = r_str_Parame & "null, "  'TABLOG_VALNUM02
   r_str_Parame = r_str_Parame & "null, "  'TABLOG_VALNUM03
   r_str_Parame = r_str_Parame & "null, "  'TABLOG_VALNUM04
   r_str_Parame = r_str_Parame & "null, "  'TABLOG_VALNUM05
   r_str_Parame = r_str_Parame & "null, "  'TABLOG_VALNUM06
   r_str_Parame = r_str_Parame & "null, "  'TABLOG_VALNUM07
   r_str_Parame = r_str_Parame & "null, "  'TABLOG_VALNUM08
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_NombPC & "', "
   r_str_Parame = r_str_Parame & "'" & UCase(App.EXEName) & "', "
   r_str_Parame = r_str_Parame & "'" & modgen_g_str_CodSuc & "') "
            
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 2) Then
      Exit Sub
   End If
   
   MsgBox "Se completo la impresión con éxito.", vbInformation, modgen_g_str_NomPlt
End Sub

Private Sub Imprimir(Path As String)
Dim Free_File As Integer
Dim Datos As String
Dim pos As Integer
Dim L As String
Dim Palabra As String
      
    ' número de archivo libre
    Free_File = FreeFile
      
    ' abre el archivo para leerlo
    Open Path For Input As Free_File
      
    ' Almacena los datos del archivo en la variable
    Datos = Input(LOF(Free_File), Free_File)
    ' cierra el archivo
    Close Free_File
      
    Do While Len(Datos) > 0
          
        pos = InStr(Datos, vbCrLf)
        If pos = 0 Then
            L = Datos
            Datos = ""
        Else
            ' linea
            L = Left$(Datos, pos - 1)
              
            Datos = Mid$(Datos, pos + 2)
        End If
      
    ' palabras
    Do While Len(L) > 0
        ' posición para extraer la palabra
        pos = InStr(L, " ")
        If pos = 0 Then
            Palabra = L
            L = ""
        Else
            Palabra = Left$(L, pos)
            L = Mid$(L, pos + 1)
        End If
      
    ' verifica que no se pase del ancho de la hoja
    If (Printer.CurrentX + Printer.TextWidth(Palabra)) <= Printer.ScaleWidth Then
        ' imprime la palabra
        Printer.Print Palabra;
    ' si no imprime en la siguiente linea
    Else
        Printer.Print
        ' verifica que no se pase del alto de la hoja
        If (Printer.CurrentY + Printer.Font.Size) > Printer.ScaleHeight Then
            ' nueva hoja
            Printer.NewPage
        End If
        ' imprime la palabra
        Printer.Print Palabra;
    End If
    Loop
    Printer.Print
    Loop
          
    ' Fin. Manda a imprimir
    Printer.EndDoc
End Sub

Private Function fs_NumLetra(ByVal p_Numero As Double) As String
   Select Case p_Numero
       Case 0: fs_NumLetra = "CERO"
       Case 1: fs_NumLetra = "UNO"
       Case 2: fs_NumLetra = "DOS"
       Case 3: fs_NumLetra = "TRES"
       Case 4: fs_NumLetra = "CUATRO"
       Case 5: fs_NumLetra = "CINCO"
       Case 6: fs_NumLetra = "SEIS"
       Case 7: fs_NumLetra = "SIETE"
       Case 8: fs_NumLetra = "OCHO"
       Case 9: fs_NumLetra = "NUEVE"
       Case 10: fs_NumLetra = "DIEZ"
       Case 11: fs_NumLetra = "ONCE"
       Case 12: fs_NumLetra = "DOCE"
       Case 13: fs_NumLetra = "TRECE"
       Case 14: fs_NumLetra = "CATORCE"
       Case 15: fs_NumLetra = "QUINCE"
       Case Is < 20: fs_NumLetra = "DIECI" & fs_NumLetra(p_Numero - 10)
       Case 20: fs_NumLetra = "VEINTE"
       Case Is < 30: fs_NumLetra = "VEINTI" & fs_NumLetra(p_Numero - 20)
       Case 30: fs_NumLetra = "TREINTA"
       Case 40: fs_NumLetra = "CUARENTA"
       Case 50: fs_NumLetra = "CINCUENTA"
       Case 60: fs_NumLetra = "SESENTA"
       Case 70: fs_NumLetra = "SETENTA"
       Case 80: fs_NumLetra = "OCHENTA"
       Case 90: fs_NumLetra = "NOVENTA"
       Case Is < 100: fs_NumLetra = fs_NumLetra(Int(p_Numero \ 10) * 10) & " Y " & fs_NumLetra(p_Numero Mod 10)
       Case 100: fs_NumLetra = "CIEN"
       Case Is < 200: fs_NumLetra = "CIENTO " & fs_NumLetra(p_Numero - 100)
       Case 200, 300, 400, 600, 800: fs_NumLetra = fs_NumLetra(Int(p_Numero \ 100)) & "CIENTOS"
       Case 500: fs_NumLetra = "QUINIENTOS"
       Case 700: fs_NumLetra = "SETECIENTOS"
       Case 900: fs_NumLetra = "NOVECIENTOS"
       Case Is < 1000: fs_NumLetra = fs_NumLetra(Int(p_Numero \ 100) * 100) & " " & fs_NumLetra(p_Numero Mod 100)
       Case 1000: fs_NumLetra = "MIL"
       Case Is < 2000: fs_NumLetra = "MIL " & fs_NumLetra(p_Numero Mod 1000)
       Case Is < 1000000: fs_NumLetra = fs_NumLetra(Int(p_Numero \ 1000)) & " MIL"
           If p_Numero Mod 1000 Then fs_NumLetra = fs_NumLetra & " " & fs_NumLetra(p_Numero Mod 1000)
       Case 1000000: fs_NumLetra = "UN MILLON"
       Case Is < 2000000: fs_NumLetra = "UN MILLON " & fs_NumLetra(p_Numero Mod 1000000)
       Case Is < 1000000000000#: fs_NumLetra = fs_NumLetra(Int(p_Numero / 1000000)) & " MILLONES "
           If (p_Numero - Int(p_Numero / 1000000) * 1000000) Then fs_NumLetra = fs_NumLetra & " " & fs_NumLetra(p_Numero - Int(p_Numero / 1000000) * 1000000)
       Case 1000000000000#: fs_NumLetra = "UN BILLON"
       Case Is < 2000000000000#: fs_NumLetra = "UN BILLON " & fs_NumLetra(p_Numero - Int(p_Numero / 1000000000000#) * 1000000000000#)
       Case Else: fs_NumLetra = fs_NumLetra(Int(p_Numero / 1000000000000#)) & " BILLONES"
           If (p_Numero - Int(p_Numero / 1000000000000#) * 1000000000000#) Then fs_NumLetra = fs_NumLetra & " " & fs_NumLetra(p_Numero - Int(p_Numero / 1000000000000#) * 1000000000000#)
   End Select
End Function

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Public Sub fs_NumeroLetra()
Dim r_str_CadAux2 As String
Dim r_str_CadAux As String

   r_str_CadAux2 = Format(CDbl(pnl_Import.Caption), "###,###,##0.00")
   r_str_CadAux = fs_NumLetra(Left(r_str_CadAux2, Len(r_str_CadAux2) - 3))
   r_str_CadAux = r_str_CadAux & " CON " & Right(Format(CDbl(r_str_CadAux2), "###,###,##0.00"), 2) & "/100"
   pnl_ImpLet.Caption = r_str_CadAux
   txt_CodOrigen.Visible = False
End Sub

Private Sub ipp_FecChq_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(txt_NomDe)
   End If
End Sub

Private Sub txt_NomDe_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(cmd_Print)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-/\ ")
   End If
End Sub

Private Sub txt_NumChq_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
       Call gs_SetFocus(ipp_FecChq)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
   End If
End Sub

