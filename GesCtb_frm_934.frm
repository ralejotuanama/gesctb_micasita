VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Pruebas_01 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_ExpExc 
      Height          =   585
      Left            =   1680
      Picture         =   "GesCtb_frm_934.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exportar a Excel"
      Top             =   1080
      Width           =   585
   End
   Begin EditLib.fpDateTime ipp_FecIni 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   1365
      _Version        =   196608
      _ExtentX        =   2408
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
   Begin EditLib.fpDateTime ipp_FecFin 
      Height          =   315
      Left            =   2220
      TabIndex        =   2
      Top             =   480
      Width           =   1365
      _Version        =   196608
      _ExtentX        =   2408
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
End
Attribute VB_Name = "frm_Pruebas_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ExpExc_Click()
Call fs_GenExcRes
End Sub

Private Sub fs_GenExcRes()

Dim r_obj_Excel      As Excel.Application
Dim r_int_Contad     As Integer
Dim r_int_NroFil     As Integer
Dim r_int_NroFil2     As Integer
Dim r_int_NroFil_nuevo     As Integer
Dim r_int_NroFil_nuevo2     As Integer
Dim r_int_NoFlLi     As Integer
Dim r_int_TotReg     As Integer
Dim r_int_Col2 As Integer
Dim r_int_Fil2 As Integer
Dim r_int_NroFil_ As Integer
Dim r_int_NroFil_aux As Integer

Dim r_int_NroFilx As Integer
Dim r_int_NroFily As Integer

Dim r_int_corre  As Integer
Dim r_int_corre2  As Integer
Dim r_int_corre_ As Integer

Dim r_rst_Prindet      As ADODB.Recordset
Dim r_rst_Prindet2     As ADODB.Recordset
Dim r_rst_Prindet3     As ADODB.Recordset
Dim r_rst_Prindet4     As ADODB.Recordset
Dim r_rst_Prindet5     As ADODB.Recordset
Dim r_rst_Prindet6     As ADODB.Recordset
Dim r_rst_Prindet7     As ADODB.Recordset
Dim r_rst_Prindet8     As ADODB.Recordset
Dim r_str_Paramedet    As String
Dim r_str_Paramedet2   As String
Dim r_str_Paramedet3   As String
Dim r_str_Paramedet4   As String
Dim r_str_Paramedet5   As String
Dim r_str_Paramedet6   As String
Dim r_str_Paramedet7   As String
Dim r_str_Paramedet8   As String
Dim var, mes, anio As String
Dim sum1, sum2, sum3, sum4, sum5, sum6, sum7, sum8, sum9, sum10, sum11, sum12 As Double
Dim sum13, sum14, sum15, sum16, sum17, sum18, sum19, sum20, sum21, sum22, sum23, sum24 As Double
Dim su1, su2, su3, su4, su5, su6 As Double
Dim fechaactual As Date
fechaatual = date


 var = Format(Now(), "yyyymmdd")
 mes = Month(Now)
 anio = Year(Now)

 
 
 



   Dim r_str_FecIni  As String
   Dim r_str_FecFin  As String



  

   r_str_FecIni = Format(ipp_FecIni.Text, "yyyymmdd")
   r_str_FecFin = Format(ipp_FecFin.Text, "yyyymmdd")

   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
'   r_obj_Excel.Visible = True
   
   
   With r_obj_Excel.Sheets(1)

        r_str_Paramedet7 = ""
        r_str_Paramedet7 = r_str_Paramedet7 + " select  sum(Round(sum(A.MAEDPF_SALCAP), 2)) As SALDOTOTAL "
        r_str_Paramedet7 = r_str_Paramedet7 + " FROM CNTBL_MAEDPF A INNER JOIN MNT_PARDES B ON A.MAEDPF_CODENT_DES = B.PARDES_CODITE AND B.PARDES_CODGRP = 122 "
        r_str_Paramedet7 = r_str_Paramedet7 + " INNER JOIN MNT_PARDES C ON A.MAEDPF_CODENT_ORI = C.PARDES_CODITE AND C.PARDES_CODGRP = 122 "
        r_str_Paramedet7 = r_str_Paramedet7 + " INNER JOIN MNT_PARDES D ON A.MAEDPF_CODMON = D.PARDES_CODITE AND D.PARDES_CODGRP = 204  AND  A.MAEDPF_SITUAC = 1 "
        r_str_Paramedet7 = r_str_Paramedet7 + " AND A.MAEDPF_SITDPF <> 3 AND A.MAEDPF_FECAPE BETWEEN " & r_str_FecIni & " And " & r_str_FecFin & ""
        r_str_Paramedet7 = r_str_Paramedet7 + " GROUP BY  CASE "
        r_str_Paramedet7 = r_str_Paramedet7 + " WHEN A.MAEDPF_PLADIA <= 7 THEN '1-7' "
        r_str_Paramedet7 = r_str_Paramedet7 + " WHEN A.MAEDPF_PLADIA > 7 AND A.MAEDPF_PLADIA <= 15  THEN '8-15' "
        r_str_Paramedet7 = r_str_Paramedet7 + " WHEN A.MAEDPF_PLADIA > 15 AND A.MAEDPF_PLADIA <= 30  THEN '16-30' "
        r_str_Paramedet7 = r_str_Paramedet7 + " WHEN A.MAEDPF_PLADIA > 30 AND A.MAEDPF_PLADIA <= 60  THEN '31-60' "
        r_str_Paramedet7 = r_str_Paramedet7 + " Else '180-365' END , B.PARDES_DESCRI "
                  
        If Not gf_EjecutaSQL(r_str_Paramedet7, r_rst_Prindet7, 3) Then
           Exit Sub
        End If
        .Cells(1, 6) = "'" & r_rst_Prindet7!SALDOTOTAL

        .Cells(1, 2) = "Tipo de Cambio"
        .Cells(2, 2) = "Patrimonio Efectivo"
    
        r_str_Paramedet5 = ""
        r_str_Paramedet5 = "SELECT DISTINCT(P.TIPCAM_COMPRA) AS TIPCOMPRA FROM OPE_TIPCAM P WHERE P.TIPCAM_FECDIA = " & var & ""
     
        If Not gf_EjecutaSQL(r_str_Paramedet5, r_rst_Prindet5, 3) Then
           Exit Sub
        End If
        .Cells(1, 3) = "'" & r_rst_Prindet5!TIPCOMPRA
        .Cells(1, 3).HorizontalAlignment = xlHAlignCenter
         
        r_str_Paramedet6 = ""
        r_str_Paramedet6 = "SELECT C.CONLIM_PATEFE FROM CTB_CONLIM C WHERE C.CONLIM_CODANO = " & anio & " AND C.CONLIM_CODMES = " & mes & ""
     
        If Not gf_EjecutaSQL(r_str_Paramedet6, r_rst_Prindet6, 3) Then
           Exit Sub
        End If
        .Cells(2, 3) = r_rst_Prindet6!CONLIM_PATEFE
        .Cells(2, 3).NumberFormat = "###,###,##0.00"
        .Cells(2, 3).HorizontalAlignment = xlHAlignCenter
         
        .Cells(2, 4) = (r_rst_Prindet6!CONLIM_PATEFE * 30) / 100
        .Cells(2, 4).NumberFormat = "###,###,##0.00"
        .Cells(2, 4).HorizontalAlignment = xlHAlignCenter
        
        .Cells(4, 4) = "REPORTE DE DÉPOSITOS A PLAZOS Y SALDOS"
        .Range(.Cells(4, 1), .Cells(4, 13)).Merge
        .Range("B4:M4").HorizontalAlignment = xlHAlignCenter
      
        .Range(.Cells(4, 1), .Cells(4, 13)).Font.Name = "Calibri"
        .Range(.Cells(4, 1), .Cells(4, 13)).Font.Size = 12
        .Range(.Cells(4, 1), .Cells(4, 13)).Font.Bold = True
        

        .Range(.Cells(3, 1), .Cells(3, 12)).HorizontalAlignment = xlHAlignCenter
      
        .Cells(1, 5) = "Datos Externos"
        .Cells(2, 5) = "Datos Externos"
         
        .Columns("A").ColumnWidth = 40

        .Columns("B").ColumnWidth = 30
        Columns("B").HorizontalAlignment = xlHAlignCenter
        .Columns("C").ColumnWidth = 20
        Columns("C").HorizontalAlignment = xlHAlignCenter
        .Columns("D").ColumnWidth = 15
        Columns("D").HorizontalAlignment = xlHAlignCenter
        .Columns("E").ColumnWidth = 15
        Columns("E").HorizontalAlignment = xlHAlignCenter
        .Columns("F").ColumnWidth = 15
        Columns("F").HorizontalAlignment = xlHAlignCenter
        .Columns("G").ColumnWidth = 15
        Columns("G").HorizontalAlignment = xlHAlignCenter
        .Columns("H").ColumnWidth = 15
        Columns("H").HorizontalAlignment = xlHAlignCenter
        .Columns("I").ColumnWidth = 20
        Columns("I").HorizontalAlignment = xlHAlignCenter
        .Columns("J").ColumnWidth = 15
        Columns("J").HorizontalAlignment = xlHAlignCenter
        .Columns("L").ColumnWidth = 15
        Columns("L").HorizontalAlignment = xlHAlignCenter
        .Columns("M").ColumnWidth = 15
        Columns("M").HorizontalAlignment = xlHAlignCenter
        
        .Cells(6, 6) = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
        .Cells(6, 6).Font.Bold = True
        .Cells(6, 6).HorizontalAlignment = xlHAlignCenter
        
        .Cells(8, 2) = "OPERACIÓN DEL DÍA"
        .Cells(8, 21).Font.Bold = True
        
         .Cells(9, 2) = "INSTITUCIÓN"
         .Cells(9, 2).Interior.Color = RGB(3, 213, 86)
         .Cells(9, 3) = "PLAZO DIAS"
         .Cells(9, 3).Interior.Color = RGB(3, 213, 86)
         .Cells(9, 4) = "TASA %"
         .Cells(9, 4).Interior.Color = RGB(3, 213, 86)
         .Cells(9, 5) = "MONEDA"
         .Cells(9, 5).Interior.Color = RGB(3, 213, 86)
         .Cells(9, 6) = "CAPITAL"
         .Cells(9, 6).Interior.Color = RGB(3, 213, 86)
         .Cells(9, 7) = "F.APERTURA"
         .Cells(9, 7).Interior.Color = RGB(3, 213, 86)
         .Cells(9, 8) = "F.VENCIMIENTO"
         .Cells(9, 8).Interior.Color = RGB(3, 213, 86)
         .Cells(9, 9) = "RENDIMIENTO"
         .Cells(9, 9).Interior.Color = RGB(3, 213, 86)
         .Cells(9, 10) = "Nueva Tasa"
         .Cells(9, 10).Interior.Color = RGB(243, 222, 13)
         .Cells(9, 11) = "Nuevo Monto"
         .Cells(9, 11).Interior.Color = RGB(243, 222, 13)
         .Cells(9, 12) = "Nuevo Plazo"
         .Cells(9, 12).Interior.Color = RGB(243, 222, 13)
         .Cells(9, 13) = "Estado"
         .Cells(9, 13).Interior.Color = RGB(243, 222, 13)
         
         .Range(.Cells(9, 2), .Cells(9, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(9, 2), .Cells(9, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(9, 2), .Cells(9, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(9, 2), .Cells(9, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(9, 2), .Cells(9, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
         r_str_Paramedet8 = ""
       

       
      r_str_Paramedet8 = r_str_Paramedet8 + "    SELECT LPAD(A.MAEDPF_NUMCTA,8,'0') MAEDPF_NUMCTA, TRIM(B.PARDES_DESCRI) AS ENTIDAD_DEST, A.MAEDPF_NUMREF, A.MAEDPF_PLADIA, "
     r_str_Paramedet8 = r_str_Paramedet8 + "   A.MAEDPF_TASINT, A.MAEDPF_CODMON, A.MAEDPF_SALCAP, A.MAEDPF_INTAJU, "
      r_str_Paramedet8 = r_str_Paramedet8 + "   A.MAEDPF_INTCAP, TRIM(C.PARDES_DESCRI) AS ENTIDAD_ORIG, TRIM(D.PARDES_DESCRI) AS MONEDA, "
     r_str_Paramedet8 = r_str_Paramedet8 + "    A.MAEDPF_FECAPE, TO_CHAR(TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA,'yyyymmdd') AS FEC_VCTO, "
       r_str_Paramedet8 = r_str_Paramedet8 + "  DECODE(A.MAEDPF_TIPDPF,1, "
       r_str_Paramedet8 = r_str_Paramedet8 + "   CASE "
      r_str_Paramedet8 = r_str_Paramedet8 + "   WHEN (TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA)-1 = TO_DATE(SYSDATE,'DD/MM/YY') THEN 'POR VENCER' "
      r_str_Paramedet8 = r_str_Paramedet8 + "  WHEN (TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA) <= TO_DATE(SYSDATE,'DD/MM/YY') THEN 'VENCIDO' "
       r_str_Paramedet8 = r_str_Paramedet8 + "  WHEN TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA > TO_DATE(SYSDATE,'DD/MM/YY') THEN 'VIGENTE' "
       r_str_Paramedet8 = r_str_Paramedet8 + "  END, 'CERRADO') AS NOM_SITUAC, "
       r_str_Paramedet8 = r_str_Paramedet8 + "  DECODE(A.MAEDPF_TIPDPF,1, "
       r_str_Paramedet8 = r_str_Paramedet8 + "  CASE "
       r_str_Paramedet8 = r_str_Paramedet8 + "  WHEN (TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA) <= TO_DATE(SYSDATE,'DD/MM/YY') THEN 2 "
       r_str_Paramedet8 = r_str_Paramedet8 + "   WHEN TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA > TO_DATE(SYSDATE,'DD/MM/YY') THEN 1  END,"
       r_str_Paramedet8 = r_str_Paramedet8 + "   A.MAEDPF_SITDPF) AS COD_SITUAC, A.MAEDPF_NUMCTA_REF, MAEDPF_CODENT_DES, MAEDPF_CODENT_ORI, MAEDPF_TIPDPF "
      r_str_Paramedet8 = r_str_Paramedet8 + "   FROM CNTBL_MAEDPF A "
     r_str_Paramedet8 = r_str_Paramedet8 + "   INNER JOIN MNT_PARDES B ON A.MAEDPF_CODENT_DES = B.PARDES_CODITE AND B.PARDES_CODGRP = 122 "
       r_str_Paramedet8 = r_str_Paramedet8 + "  INNER JOIN MNT_PARDES C ON A.MAEDPF_CODENT_ORI = C.PARDES_CODITE AND C.PARDES_CODGRP = 122 "
      r_str_Paramedet8 = r_str_Paramedet8 + "  INNER JOIN MNT_PARDES D ON A.MAEDPF_CODMON = D.PARDES_CODITE AND D.PARDES_CODGRP = 204 "
      r_str_Paramedet8 = r_str_Paramedet8 + "    WHERE   TO_CHAR(TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA,'yyyymmdd') =  " & var & " AND  A.MAEDPF_SITUAC = 1 "
      
       
       
                  If Not gf_EjecutaSQL(r_str_Paramedet8, r_rst_Prindet8, 3) Then
                     Exit Sub
                  End If
     
       'r_obj_Excel.Visible = True
                   
                    r_int_corre_ = 1
                    r_int_NroFil_ = 10
    
                              r_rst_Prindet8.MoveFirst
                              Do While Not r_rst_Prindet8.EOF
                                .Cells(r_int_NroFil_, 2) = r_rst_Prindet8!ENTIDAD_DEST
                                .Cells(r_int_NroFil_, 3) = r_rst_Prindet8!MAEDPF_PLADIA
                                .Cells(r_int_NroFil_, 4) = r_rst_Prindet8!MAEDPF_TASINT
                                .Cells(r_int_NroFil_, 5) = r_rst_Prindet8!Moneda
                                .Cells(r_int_NroFil_, 6) = r_rst_Prindet8!MAEDPF_SALCAP
                                .Cells(r_int_NroFil_, 6).NumberFormat = "###,###,##0.00"

                                .Cells(r_int_NroFil_, 7) = Mid(r_rst_Prindet8!MAEDPF_FECAPE, 1, 4) & "-" & Mid(r_rst_Prindet8!MAEDPF_FECAPE, 5, 2) & "-" & Mid(r_rst_Prindet8!MAEDPF_FECAPE, 7, 2)
                                .Cells(r_int_NroFil_, 8) = Mid(r_rst_Prindet8!FEC_VCTO, 1, 4) & "-" & Mid(r_rst_Prindet8!FEC_VCTO, 5, 2) & "-" & Mid(r_rst_Prindet8!FEC_VCTO, 7, 2)

                               .Cells(r_int_NroFil_, 9) = r_rst_Prindet8!MAEDPF_INTCAP
                               .Cells(r_int_NroFil_, 9).NumberFormat = "###,###,##0.00"
                                
                             r_int_corre_ = r_int_corre_ + 1
                             r_int_NroFil_ = r_int_NroFil_ + 1
                             r_rst_Prindet8.MoveNext
                             DoEvents
                             Loop
                             
        r_int_NroFil_aux = r_int_NroFil_ + 2
           
        .Cells(r_int_NroFil_aux, 1) = "VENCIMIENTOS DEL MES EN"
        .Cells(r_int_NroFil_aux, 1).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 2) = "ENE-2020"
        .Cells(r_int_NroFil_aux, 2).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 3) = "FEB-2020"
        .Cells(r_int_NroFil_aux, 3).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 4) = "MAR-2020"
        .Cells(r_int_NroFil_aux, 4).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 5) = "ABR-2020"
        .Cells(r_int_NroFil_aux, 5).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 6) = "MAY-2020"
        .Cells(r_int_NroFil_aux, 6).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 7) = "JUN-2020"
        .Cells(r_int_NroFil_aux, 7).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 8) = "JUL-2020"
        .Cells(r_int_NroFil_aux, 8).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 9) = "AGO-2020"
        .Cells(r_int_NroFil_aux, 9).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 10) = "SET-2020"
        .Cells(r_int_NroFil_aux, 10).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 11) = "OCT-2020"
        .Cells(r_int_NroFil_aux, 11).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 12) = "NOV-2020"
        .Cells(r_int_NroFil_aux, 12).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 13) = "DIC-2020"
        .Cells(r_int_NroFil_aux, 13).Font.Bold = True
        
         .Cells(r_int_NroFil_aux, 14) = "ENE-2021"
        .Cells(r_int_NroFil_aux, 14).Font.Bold = True
        
         .Cells(r_int_NroFil_aux, 15) = "FEB-2021"
        .Cells(r_int_NroFil_aux, 15).Font.Bold = True
        
         .Cells(r_int_NroFil_aux, 16) = "MAR-2021"
        .Cells(r_int_NroFil_aux, 16).Font.Bold = True
        
         .Cells(r_int_NroFil_aux, 17) = "ABR-2021"
        .Cells(r_int_NroFil_aux, 17).Font.Bold = True
        
        
        .Cells(r_int_NroFil_aux, 18) = "MAY-2021"
        .Cells(r_int_NroFil_aux, 18).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 19) = "JUN-2021"
        .Cells(r_int_NroFil_aux, 19).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 20) = "JUL-2021"
        .Cells(r_int_NroFil_aux, 20).Font.Bold = True
        
        .Cells(r_int_NroFil_aux, 21) = "AGO-2021"
        .Cells(r_int_NroFil_aux, 21).Font.Bold = True
        
         .Cells(r_int_NroFil_aux, 22) = "SET-2021"
        .Cells(r_int_NroFil_aux, 22).Font.Bold = True
        
         .Cells(r_int_NroFil_aux, 23) = "OCT-2021"
        .Cells(r_int_NroFil_aux, 23).Font.Bold = True
        
         .Cells(r_int_NroFil_aux, 24) = "NOV-2021"
        .Cells(r_int_NroFil_aux, 24).Font.Bold = True
        
         .Cells(r_int_NroFil_aux, 25) = "DIC-2021"
        .Cells(r_int_NroFil_aux, 25).Font.Bold = True
        
        
        
        
        .Range(.Cells(r_int_NroFil_aux, 1), .Cells(r_int_NroFil_aux, 25)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range(.Cells(r_int_NroFil_aux, 1), .Cells(r_int_NroFil_aux, 25)).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(r_int_NroFil_aux, 1), .Cells(r_int_NroFil_aux, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(r_int_NroFil_aux, 1), .Cells(r_int_NroFil_aux, 25)).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range(.Cells(r_int_NroFil_aux, 1), .Cells(r_int_NroFil_aux, 25)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(.Cells(r_int_NroFil_aux, 1), .Cells(r_int_NroFil_aux, 25)).HorizontalAlignment = xlHAlignCenter
        
         
             g_str_Parame = ""
             g_str_Parame = g_str_Parame + " SELECT COUNT(*) AS CONT, SUBSTR(TO_CHAR(TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA,'yyyymmdd'),5,2) AS MES , B.PARDES_DESCRI , ROUND(SUM(A.MAEDPF_SALCAP),2) AS SALDO,  SUBSTR(TO_CHAR(TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA,'yyyymmdd'),1,4) AS ANIO"
             g_str_Parame = g_str_Parame + " FROM CNTBL_MAEDPF A INNER JOIN MNT_PARDES B ON A.MAEDPF_CODENT_DES = B.PARDES_CODITE AND B.PARDES_CODGRP = 122 "
             g_str_Parame = g_str_Parame + " INNER JOIN MNT_PARDES C ON A.MAEDPF_CODENT_ORI = C.PARDES_CODITE AND C.PARDES_CODGRP = 122  "
             g_str_Parame = g_str_Parame + "  INNER JOIN MNT_PARDES D ON A.MAEDPF_CODMON = D.PARDES_CODITE AND D.PARDES_CODGRP = 204 "
             g_str_Parame = g_str_Parame + " WHERE A.MAEDPF_SITUAC = 1 "
             g_str_Parame = g_str_Parame + " AND A.MAEDPF_SITDPF <> 3 AND A.MAEDPF_FECAPE BETWEEN " & r_str_FecIni & " And " & r_str_FecFin & ""
             g_str_Parame = g_str_Parame + " GROUP BY  SUBSTR(TO_CHAR(TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA,'yyyymmdd'),5,2), B.PARDES_DESCRI, SUBSTR(TO_CHAR(TO_DATE(A.MAEDPF_FECAPE, 'yyyymmdd') + A.MAEDPF_PLADIA,'yyyymmdd'),1,4) "
             g_str_Parame = g_str_Parame + "ORDER BY ANIO,MES,B.PARDES_DESCRI "
    
             If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
                Exit Sub
             End If
         
'           r_obj_Excel.Visible = True
           
            g_rst_Princi.MoveFirst
            Do While Not g_rst_Princi.EOF
                 'entidades bancarias
                 r_str_Paramedet = ""
                 r_str_Paramedet = r_str_Paramedet + " SELECT DISTINCT B.PARDES_DESCRI "
                 r_str_Paramedet = r_str_Paramedet + " FROM CNTBL_MAEDPF A INNER JOIN MNT_PARDES B ON A.MAEDPF_CODENT_DES = B.PARDES_CODITE AND B.PARDES_CODGRP = 122 "
                 r_str_Paramedet = r_str_Paramedet + " INNER JOIN MNT_PARDES C ON A.MAEDPF_CODENT_ORI = C.PARDES_CODITE AND C.PARDES_CODGRP = 122 "
                 r_str_Paramedet = r_str_Paramedet + " INNER JOIN MNT_PARDES D ON A.MAEDPF_CODMON = D.PARDES_CODITE AND D.PARDES_CODGRP = 204 "
                 r_str_Paramedet = r_str_Paramedet + " AND  A.MAEDPF_SITUAC = 1 "
                 r_str_Paramedet = r_str_Paramedet + " AND A.MAEDPF_SITDPF <> 3 AND A.MAEDPF_FECAPE BETWEEN " & r_str_FecIni & " And " & r_str_FecFin & ""
                 r_str_Paramedet = r_str_Paramedet + " ORDER BY B.PARDES_DESCRI "
     
                 If Not gf_EjecutaSQL(r_str_Paramedet, r_rst_Prindet, 3) Then
                     Exit Sub
                  End If
     
                    r_int_NroFil = r_int_NroFil_aux + 1
                    r_int_corre = 1
    
                          r_rst_Prindet.MoveFirst
                           Do While Not r_rst_Prindet.EOF
                            .Cells(r_int_NroFil, 1) = r_rst_Prindet!PARDES_DESCRI
                             If Trim(g_rst_Princi!PARDES_DESCRI) = Trim(r_rst_Prindet!PARDES_DESCRI) Then
                             
                                If g_rst_Princi!mes = "01" And g_rst_Princi!anio = "2020" Then
                                  .Cells(r_int_NroFil, 2) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                  .Cells(r_int_NroFil, 2).NumberFormat = "###,###,##0.00"
                                  sum1 = sum1 + CDbl(.Cells(r_int_NroFil, 2))
                                  
                                ElseIf g_rst_Princi!mes = "02" And g_rst_Princi!anio = "2020" Then
                                 .Cells(r_int_NroFil, 3) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 3).NumberFormat = "###,###,##0.00"
                                  sum2 = sum2 + CDbl(.Cells(r_int_NroFil, 3))
                                  
                                ElseIf g_rst_Princi!mes = "03" And g_rst_Princi!anio = "2020" Then
                                 .Cells(r_int_NroFil, 4) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 4).NumberFormat = "###,###,##0.00"
                                  sum3 = sum3 + CDbl(.Cells(r_int_NroFil, 4))
                                    
                                  
                                ElseIf g_rst_Princi!mes = "04" And g_rst_Princi!anio = "2020" Then
                                 .Cells(r_int_NroFil, 5) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 5).NumberFormat = "###,###,##0.00"
                                 sum4 = sum4 + CDbl(.Cells(r_int_NroFil, 5))
                                 
                                ElseIf g_rst_Princi!mes = "05" And g_rst_Princi!anio = "2020" Then
                                 .Cells(r_int_NroFil, 6) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 6).NumberFormat = "###,###,##0.00"
                                 sum5 = sum5 + CDbl(.Cells(r_int_NroFil, 6))
                                 
                                ElseIf g_rst_Princi!mes = "06" And g_rst_Princi!anio = "2020" Then
                                 .Cells(r_int_NroFil, 7) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 7).NumberFormat = "###,###,##0.00"
                                 sum6 = sum6 + CDbl(.Cells(r_int_NroFil, 7))
                                 
                                ElseIf g_rst_Princi!mes = "07" And g_rst_Princi!anio = "2020" Then
                                 .Cells(r_int_NroFil, 8) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 8).NumberFormat = "###,###,##0.00"
                                sum7 = sum7 + CDbl(.Cells(r_int_NroFil, 8))
                                
                                ElseIf g_rst_Princi!mes = "08" And g_rst_Princi!anio = "2020" Then
                                  .Cells(r_int_NroFil, 9) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                  .Cells(r_int_NroFil, 9).NumberFormat = "###,###,##0.00"
                                 sum8 = sum8 + CDbl(.Cells(r_int_NroFil, 9))
                                 
                                ElseIf g_rst_Princi!mes = "09" And g_rst_Princi!anio = "2020" Then
                                  .Cells(r_int_NroFil, 10) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                  .Cells(r_int_NroFil, 10).NumberFormat = "###,###,##0.00"
                                 sum9 = sum9 + CDbl(.Cells(r_int_NroFil, 10))
                                 
                                ElseIf g_rst_Princi!mes = "10" And g_rst_Princi!anio = "2020" Then
                                 .Cells(r_int_NroFil, 11) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 11).NumberFormat = "###,###,##0.00"
                                 sum10 = sum10 + CDbl(.Cells(r_int_NroFil, 11))
                                 
                                ElseIf g_rst_Princi!mes = "11" And g_rst_Princi!anio = "2020" Then
                                 .Cells(r_int_NroFil, 12) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 12).NumberFormat = "###,###,##0.00"
                                 sum11 = sum11 + CDbl(.Cells(r_int_NroFil, 12))
                                 
                                ElseIf g_rst_Princi!mes = "12" And g_rst_Princi!anio = "2020" Then
                                 .Cells(r_int_NroFil, 13) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 13).NumberFormat = "###,###,##0.00"
                                 sum12 = sum12 + CDbl(.Cells(r_int_NroFil, 13))
                                 
                                  ElseIf g_rst_Princi!mes = "01" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 14) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 14).NumberFormat = "###,###,##0.00"
                                 sum13 = sum13 + CDbl(.Cells(r_int_NroFil, 14))
                                 
                                   ElseIf g_rst_Princi!mes = "02" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 15) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 15).NumberFormat = "###,###,##0.00"
                                 sum14 = sum14 + CDbl(.Cells(r_int_NroFil, 15))
                                 
                                 
                                  ElseIf g_rst_Princi!mes = "03" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 16) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 16).NumberFormat = "###,###,##0.00"
                                 sum15 = sum15 + CDbl(.Cells(r_int_NroFil, 16))
                                 
                                 
                                   ElseIf g_rst_Princi!mes = "04" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 17) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 17).NumberFormat = "###,###,##0.00"
                                 sum16 = sum16 + CDbl(.Cells(r_int_NroFil, 17))
                                 
                                 
                                 
                                   ElseIf g_rst_Princi!mes = "05" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 18) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 18).NumberFormat = "###,###,##0.00"
                                 sum17 = sum17 + CDbl(.Cells(r_int_NroFil, 18))
                                 
                                 
                                   ElseIf g_rst_Princi!mes = "06" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 19) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 19).NumberFormat = "###,###,##0.00"
                                 sum18 = sum18 + CDbl(.Cells(r_int_NroFil, 19))
                                 
                                 
                                   ElseIf g_rst_Princi!mes = "07" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 20) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 20).NumberFormat = "###,###,##0.00"
                                 sum19 = sum19 + CDbl(.Cells(r_int_NroFil, 20))
                                 
                                 
                                 
                                   ElseIf g_rst_Princi!mes = "08" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 21) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 21).NumberFormat = "###,###,##0.00"
                                 sum20 = sum20 + CDbl(.Cells(r_int_NroFil, 21))
                                 
                                 
                                   ElseIf g_rst_Princi!mes = "09" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 22) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 22).NumberFormat = "###,###,##0.00"
                                 sum21 = sum21 + CDbl(.Cells(r_int_NroFil, 22))
                                 
                                 
                                   ElseIf g_rst_Princi!mes = "10" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 23) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 23).NumberFormat = "###,###,##0.00"
                                 sum22 = sum22 + CDbl(.Cells(r_int_NroFil, 23))
                                 
                                 
                                 
                                 
                                   ElseIf g_rst_Princi!mes = "11" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 24) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 24).NumberFormat = "###,###,##0.00"
                                 sum23 = sum23 + CDbl(.Cells(r_int_NroFil, 24))
                                 
                                 
                                 
                                   ElseIf g_rst_Princi!mes = "12" And g_rst_Princi!anio = "2021" Then
                                 .Cells(r_int_NroFil, 25) = IIf(g_rst_Princi!SALDO > 0, g_rst_Princi!SALDO, 0)
                                 .Cells(r_int_NroFil, 25).NumberFormat = "###,###,##0.00"
                                 sum24 = sum24 + CDbl(.Cells(r_int_NroFil, 25))
                                 
                                 
                                 
                                 
                                 
                                 
                                 
                                 
                                End If
                             End If
                    
                             r_int_corre = r_int_corre + 1
                             r_int_NroFil = r_int_NroFil + 1
                             r_rst_Prindet.MoveNext
                             DoEvents
                             Loop
                             
                           r_int_NroFilx = r_int_NroFil
                             
                     .Cells(r_int_NroFilx, 1) = "Total General"
                     .Cells(r_int_NroFilx, 2) = sum1
                     .Cells(r_int_NroFilx, 2).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 2).Font.Bold = True
                     .Cells(r_int_NroFilx, 3) = sum2
                     .Cells(r_int_NroFilx, 3).Font.Bold = True
                     .Cells(r_int_NroFilx, 3).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 4) = sum3
                     .Cells(r_int_NroFilx, 4).Font.Bold = True
                     .Cells(r_int_NroFilx, 4).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 5) = sum4
                     .Cells(r_int_NroFilx, 5).Font.Bold = True
                     .Cells(r_int_NroFilx, 5).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 6) = sum5
                     .Cells(r_int_NroFilx, 6).Font.Bold = True
                     .Cells(r_int_NroFilx, 6).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 7) = sum6
                     .Cells(r_int_NroFilx, 7).Font.Bold = True
                     .Cells(r_int_NroFilx, 7).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 8) = sum7
                     .Cells(r_int_NroFilx, 8).Font.Bold = True
                     .Cells(r_int_NroFilx, 8).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 9) = sum8
                     .Cells(r_int_NroFilx, 9).Font.Bold = True
                     .Cells(r_int_NroFilx, 9).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 10) = sum9
                     .Cells(r_int_NroFilx, 10).Font.Bold = True
                     .Cells(r_int_NroFilx, 10).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 11) = sum10
                     .Cells(r_int_NroFilx, 11).Font.Bold = True
                     .Cells(r_int_NroFilx, 11).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 12) = sum11
                     .Cells(r_int_NroFilx, 12).Font.Bold = True
                     .Cells(r_int_NroFilx, 12).NumberFormat = "###,###,##0.00"
                     .Cells(r_int_NroFilx, 13) = sum12
                     .Cells(r_int_NroFilx, 13).Font.Bold = True
                     .Cells(r_int_NroFilx, 13).NumberFormat = "###,###,##0.00"
                      .Cells(r_int_NroFilx, 14) = sum13
                     .Cells(r_int_NroFilx, 14).Font.Bold = True
                     .Cells(r_int_NroFilx, 14).NumberFormat = "###,###,##0.00"
                     
                      .Cells(r_int_NroFilx, 15) = sum14
                     .Cells(r_int_NroFilx, 15).Font.Bold = True
                     .Cells(r_int_NroFilx, 15).NumberFormat = "###,###,##0.00"
                     
                     
                     
                      .Cells(r_int_NroFilx, 16) = sum15
                     .Cells(r_int_NroFilx, 16).Font.Bold = True
                     .Cells(r_int_NroFilx, 16).NumberFormat = "###,###,##0.00"
                     
                     
                      .Cells(r_int_NroFilx, 17) = sum16
                     .Cells(r_int_NroFilx, 17).Font.Bold = True
                     .Cells(r_int_NroFilx, 17).NumberFormat = "###,###,##0.00"
                     
                     
                      .Cells(r_int_NroFilx, 18) = sum17
                     .Cells(r_int_NroFilx, 18).Font.Bold = True
                     .Cells(r_int_NroFilx, 18).NumberFormat = "###,###,##0.00"
                     
                     
                      .Cells(r_int_NroFilx, 19) = sum18
                     .Cells(r_int_NroFilx, 19).Font.Bold = True
                     .Cells(r_int_NroFilx, 19).NumberFormat = "###,###,##0.00"
                     
                     
                       .Cells(r_int_NroFilx, 20) = sum19
                     .Cells(r_int_NroFilx, 20).Font.Bold = True
                     .Cells(r_int_NroFilx, 20).NumberFormat = "###,###,##0.00"
                     
                     
                       .Cells(r_int_NroFilx, 21) = sum20
                     .Cells(r_int_NroFilx, 21).Font.Bold = True
                     .Cells(r_int_NroFilx, 21).NumberFormat = "###,###,##0.00"
                     
                       .Cells(r_int_NroFilx, 22) = sum21
                     .Cells(r_int_NroFilx, 22).Font.Bold = True
                     .Cells(r_int_NroFilx, 22).NumberFormat = "###,###,##0.00"
                     
                       .Cells(r_int_NroFilx, 23) = sum22
                     .Cells(r_int_NroFilx, 23).Font.Bold = True
                     .Cells(r_int_NroFilx, 23).NumberFormat = "###,###,##0.00"
                     
                       .Cells(r_int_NroFilx, 24) = sum23
                     .Cells(r_int_NroFilx, 24).Font.Bold = True
                     .Cells(r_int_NroFilx, 24).NumberFormat = "###,###,##0.00"
                     
                       .Cells(r_int_NroFilx, 25) = sum24
                     .Cells(r_int_NroFilx, 25).Font.Bold = True
                     .Cells(r_int_NroFilx, 25).NumberFormat = "###,###,##0.00"
                     
                     
                     
                     
                     .Range(.Cells(r_int_NroFilx, 1), .Cells(r_int_NroFil, 25)).Borders(xlEdgeLeft).LineStyle = xlContinuous
                     .Range(.Cells(r_int_NroFilx, 1), .Cells(r_int_NroFil, 25)).Borders(xlEdgeTop).LineStyle = xlContinuous
                     .Range(.Cells(r_int_NroFilx, 1), .Cells(r_int_NroFil, 25)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                     .Range(.Cells(r_int_NroFilx, 1), .Cells(r_int_NroFil, 25)).Borders(xlEdgeRight).LineStyle = xlContinuous
                     .Range(.Cells(r_int_NroFilx, 1), .Cells(r_int_NroFil, 25)).Borders(xlInsideVertical).LineStyle = xlContinuous
                     
                     .Range(.Cells(r_int_NroFilx, 1), .Cells(r_int_NroFilx, 25)).HorizontalAlignment = xlHAlignCenter
                       
                 g_rst_Princi.MoveNext
                 DoEvents
            Loop
            
            r_int_NroFil_nuevo = r_int_NroFilx + 3


            .Range(.Cells(r_int_NroFil_nuevo, 1), .Cells(40, 8)).NumberFormat = "@"
            .Cells(r_int_NroFil_nuevo, 1) = "DPF SOLES"
            .Cells(r_int_NroFil_nuevo, 1).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 2) = "07-15"
            .Cells(r_int_NroFil_nuevo, 2).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 3) = "16-30"
            .Cells(r_int_NroFil_nuevo, 3).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 4) = "31-60"
            .Cells(r_int_NroFil_nuevo, 4).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 5) = "180-365"
            .Cells(r_int_NroFil_nuevo, 5).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 6) = "Total General"
            .Cells(r_int_NroFil_nuevo, 6).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 7) = "En S/."
            .Cells(r_int_NroFil_nuevo, 7).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 8) = "% TOTAL DPF"
            .Cells(r_int_NroFil_nuevo, 8).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 9) = "% Patrimonio Efectivo"
            .Cells(r_int_NroFil_nuevo, 9).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 10) = "Saldo en Cuenta"
            .Cells(r_int_NroFil_nuevo, 10).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 11) = "En S/."
            .Cells(r_int_NroFil_nuevo, 11).Font.Bold = True
            .Cells(r_int_NroFil_nuevo, 12) = "Holgura"
            .Cells(r_int_NroFil_nuevo, 12).Font.Bold = True
            
             .Range(.Cells(r_int_NroFil_nuevo, 1), .Cells(r_int_NroFil_nuevo, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
             .Range(.Cells(r_int_NroFil_nuevo, 1), .Cells(r_int_NroFil_nuevo, 12)).Borders(xlEdgeTop).LineStyle = xlContinuous
             .Range(.Cells(r_int_NroFil_nuevo, 1), .Cells(r_int_NroFil_nuevo, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
             .Range(.Cells(r_int_NroFil_nuevo, 1), .Cells(r_int_NroFil_nuevo, 12)).Borders(xlEdgeRight).LineStyle = xlContinuous
             .Range(.Cells(r_int_NroFil_nuevo, 1), .Cells(r_int_NroFil_nuevo, 12)).Borders(xlInsideVertical).LineStyle = xlContinuous
             .Range(.Cells(r_int_NroFil_nuevo, 1), .Cells(r_int_NroFil_nuevo, 12)).HorizontalAlignment = xlHAlignCenter
                 
             r_int_NroFil_nuevo2 = r_int_NroFil_nuevo + 1
         
            r_str_Paramedet3 = ""
            r_str_Paramedet3 = r_str_Paramedet3 + " SELECT COUNT(*) AS CONT ,"
            r_str_Paramedet3 = r_str_Paramedet3 + " CASE "
            r_str_Paramedet3 = r_str_Paramedet3 + " WHEN A.MAEDPF_PLADIA <= 7 THEN '1-7' "
            r_str_Paramedet3 = r_str_Paramedet3 + " WHEN A.MAEDPF_PLADIA > 7 AND A.MAEDPF_PLADIA <= 15  THEN '8-15' "
            r_str_Paramedet3 = r_str_Paramedet3 + " WHEN A.MAEDPF_PLADIA > 15 AND A.MAEDPF_PLADIA <= 30  THEN '16-30' "
            r_str_Paramedet3 = r_str_Paramedet3 + " WHEN A.MAEDPF_PLADIA > 30 AND A.MAEDPF_PLADIA <= 60  THEN '31-60' "
            r_str_Paramedet3 = r_str_Paramedet3 + " Else '180-365' "
            r_str_Paramedet3 = r_str_Paramedet3 + " END  AS AB ,"
            r_str_Paramedet3 = r_str_Paramedet3 + " B.PARDES_DESCRI,  "
            r_str_Paramedet3 = r_str_Paramedet3 + " ROUND(Sum(A.MAEDPF_SALCAP), 2) As SALDO "
            r_str_Paramedet3 = r_str_Paramedet3 + "  FROM CNTBL_MAEDPF A INNER JOIN MNT_PARDES B ON A.MAEDPF_CODENT_DES = B.PARDES_CODITE AND B.PARDES_CODGRP = 122 "
            r_str_Paramedet3 = r_str_Paramedet3 + " INNER JOIN MNT_PARDES C ON A.MAEDPF_CODENT_ORI = C.PARDES_CODITE AND C.PARDES_CODGRP = 122 "
            r_str_Paramedet3 = r_str_Paramedet3 + " INNER JOIN MNT_PARDES D ON A.MAEDPF_CODMON = D.PARDES_CODITE AND D.PARDES_CODGRP = 204  AND  A.MAEDPF_SITUAC = 1  AND A.MAEDPF_SITDPF <> 3 AND A.MAEDPF_FECAPE BETWEEN " & r_str_FecIni & " And " & r_str_FecFin & ""
            r_str_Paramedet3 = r_str_Paramedet3 + " GROUP BY  CASE "
            r_str_Paramedet3 = r_str_Paramedet3 + " WHEN A.MAEDPF_PLADIA <= 7 THEN '1-7' "
            r_str_Paramedet3 = r_str_Paramedet3 + " WHEN A.MAEDPF_PLADIA > 7 AND A.MAEDPF_PLADIA <= 15  THEN '8-15' "
            r_str_Paramedet3 = r_str_Paramedet3 + " WHEN A.MAEDPF_PLADIA > 15 AND A.MAEDPF_PLADIA <= 30  THEN '16-30' "
            r_str_Paramedet3 = r_str_Paramedet3 + " WHEN A.MAEDPF_PLADIA > 30 AND A.MAEDPF_PLADIA <= 60  THEN '31-60' "
            r_str_Paramedet3 = r_str_Paramedet3 + " ELSE '180-365' END , B.PARDES_DESCRI  ORDER BY AB "
            r_obj_Excel.Visible = True
            If Not gf_EjecutaSQL(r_str_Paramedet3, r_rst_Prindet3, 3) Then
               Exit Sub
            End If
            
               r_rst_Prindet3.MoveFirst
               Do While Not r_rst_Prindet3.EOF

                             r_str_Paramedet4 = ""
                             r_str_Paramedet4 = r_str_Paramedet4 + " SELECT DISTINCT B.PARDES_DESCRI "
                             r_str_Paramedet4 = r_str_Paramedet4 + " FROM CNTBL_MAEDPF A INNER JOIN MNT_PARDES B ON A.MAEDPF_CODENT_DES = B.PARDES_CODITE AND B.PARDES_CODGRP = 122 "
                             r_str_Paramedet4 = r_str_Paramedet4 + " INNER JOIN MNT_PARDES C ON A.MAEDPF_CODENT_ORI = C.PARDES_CODITE AND C.PARDES_CODGRP = 122 "
                             r_str_Paramedet4 = r_str_Paramedet4 + " INNER JOIN MNT_PARDES D ON A.MAEDPF_CODMON = D.PARDES_CODITE AND D.PARDES_CODGRP = 204 "
                             r_str_Paramedet4 = r_str_Paramedet4 + " AND  A.MAEDPF_SITUAC = 1 "
                             r_str_Paramedet4 = r_str_Paramedet4 + " AND A.MAEDPF_SITDPF <> 3 AND A.MAEDPF_FECAPE BETWEEN " & r_str_FecIni & " And " & r_str_FecFin & ""
                             r_str_Paramedet4 = r_str_Paramedet4 + " ORDER BY B.PARDES_DESCRI "
        
                             If Not gf_EjecutaSQL(r_str_Paramedet4, r_rst_Prindet4, 3) Then
                                 Exit Sub
                              End If
                              
                            r_int_NroFil2 = r_int_NroFil_nuevo2
                            r_int_corre2 = 1
             
                         r_rst_Prindet4.MoveFirst
                         Do While Not r_rst_Prindet4.EOF
                
                                  .Cells(r_int_NroFil2, 1) = r_rst_Prindet4!PARDES_DESCRI
                                  
                                If Trim(r_rst_Prindet3!PARDES_DESCRI) = Trim(r_rst_Prindet4!PARDES_DESCRI) Then
                                   
                                    If r_rst_Prindet3!AB = "7-15" Then
                                    .Cells(r_int_NroFil2, 2) = IIf(r_rst_Prindet3!SALDO > 0, r_rst_Prindet3!SALDO, 0)
                                    .Cells(r_int_NroFil2, 2).NumberFormat = "###,###,##0.00"
                                    su1 = su1 + CDbl(.Cells(r_int_NroFil2, 2))
                                  
                                    ElseIf r_rst_Prindet3!AB = "16-30" Then
                                    .Cells(r_int_NroFil2, 3) = IIf(r_rst_Prindet3!SALDO > 0, r_rst_Prindet3!SALDO, 0)
                                    .Cells(r_int_NroFil2, 3).NumberFormat = "###,###,##0.00"
                                    su2 = su2 + CDbl(.Cells(r_int_NroFil2, 3))
                                  
                                    ElseIf r_rst_Prindet3!AB = "31-60" Then
                                    .Cells(r_int_NroFil2, 4) = IIf(r_rst_Prindet3!SALDO > 0, r_rst_Prindet3!SALDO, 0)
                                    .Cells(r_int_NroFil2, 4).NumberFormat = "###,###,##0.00"
                                    su3 = su3 + CDbl(.Cells(r_int_NroFil2, 4))
                                  
                                    ElseIf r_rst_Prindet3!AB = "180-365" Then
                                     .Cells(r_int_NroFil2, 5) = IIf(r_rst_Prindet3!SALDO > 0, r_rst_Prindet3!SALDO, 0)
                                     .Cells(r_int_NroFil2, 5).NumberFormat = "###,###,##0.00"
                                     su4 = su4 + CDbl(.Cells(r_int_NroFil2, 5))
                                    End If
                                    
                                    .Cells(r_int_NroFil2, 6) = CDbl(.Cells(r_int_NroFil2, 2)) + (.Cells(r_int_NroFil2, 3)) + CDbl(.Cells(r_int_NroFil2, 4)) + CDbl(.Cells(r_int_NroFil2, 5))
                                    .Cells(r_int_NroFil2, 6).NumberFormat = "###,###,##0.00"
                                      su5 = su5 + CDbl(.Cells(r_int_NroFil2, 6))
                                      
                                    .Cells(r_int_NroFil2, 7) = .Cells(r_int_NroFil2, 2) + .Cells(r_int_NroFil2, 3) + .Cells(r_int_NroFil2, 4) + .Cells(r_int_NroFil2, 5)
                                    .Cells(r_int_NroFil2, 7).NumberFormat = "###,###,##0.00"
                                      su6 = su6 + CDbl(.Cells(r_int_NroFil2, 7))
                                   
                                    .Cells(r_int_NroFil2, 8) = "'" & CStr(Format((.Cells(r_int_NroFil2, 2) + .Cells(r_int_NroFil2, 3) + .Cells(r_int_NroFil2, 4) + .Cells(r_int_NroFil2, 5)) / (.Cells(1, 6)), "###,###,##0.00") * 100) & " %"
                                    .Cells(r_int_NroFil2, 8).NumberFormat = "###,###,##0.00"
                            
                                    .Cells(r_int_NroFil2, 9) = "'" & CStr(Format((.Cells(r_int_NroFil2, 6) / .Cells(2, 3)) * 100, "###,###,##0.00")) & " %"
                                       
                                    .Cells(r_int_NroFil2, 12) = .Cells(2, 3) - .Cells(r_int_NroFil2, 2) - .Cells(r_int_NroFil2, 3) - .Cells(r_int_NroFil2, 4) - .Cells(r_int_NroFil2, 5)
                                    .Cells(r_int_NroFil2, 12).NumberFormat = "###,###,##0.00"
                                       su12 = su12 + .Cells(r_int_NroFil2, 12)
                                End If

                                         r_int_corre2 = r_int_corre2 + 1
                                         r_int_NroFil2 = r_int_NroFil2 + 1
                                         r_rst_Prindet4.MoveNext
                          DoEvents
                          Loop
          
          
               r_rst_Prindet3.MoveNext
               DoEvents
               Loop
               
                r_int_NroFily = r_int_NroFil2
               
               Dim r_dbl_tot As Double
                  r_dbl_tot = .Cells(r_int_NroFil2, 7)
                 .Cells(r_int_NroFil2, 1) = "Total general"
                 .Cells(r_int_NroFil2, 2) = su1
                 .Cells(r_int_NroFil2, 2).NumberFormat = "###,###,##0.00"
                 .Cells(r_int_NroFil2, 3) = su2
                 .Cells(r_int_NroFil2, 3).NumberFormat = "###,###,##0.00"
                 .Cells(r_int_NroFil2, 4) = su3
                 .Cells(r_int_NroFil2, 4).NumberFormat = "###,###,##0.00"
                 .Cells(r_int_NroFil2, 5) = su4
                 .Cells(r_int_NroFil2, 5).NumberFormat = "###,###,##0.00"
                 .Cells(r_int_NroFil2, 6) = su5
                 .Cells(r_int_NroFil2, 6).NumberFormat = "###,###,##0.00"
                 .Cells(r_int_NroFil2, 7) = su6
                 .Cells(r_int_NroFil2, 7).NumberFormat = "###,###,##0.00"
                 .Cells(r_int_NroFil2, 12) = su12
                 .Cells(r_int_NroFil2, 12).NumberFormat = "###,###,##0.00"
                
                  .Range(.Cells(r_int_NroFil2, 1), .Cells(r_int_NroFil2, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
                  .Range(.Cells(r_int_NroFil2, 1), .Cells(r_int_NroFil2, 12)).Borders(xlEdgeTop).LineStyle = xlContinuous
                  .Range(.Cells(r_int_NroFil2, 1), .Cells(r_int_NroFil2, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                  .Range(.Cells(r_int_NroFil2, 1), .Cells(r_int_NroFil2, 12)).Borders(xlEdgeRight).LineStyle = xlContinuous
                  .Range(.Cells(r_int_NroFil2, 1), .Cells(r_int_NroFil2, 12)).Borders(xlInsideVertical).LineStyle = xlContinuous
                  .Range(.Cells(r_int_NroFil2, 1), .Cells(r_int_NroFil2, 12)).HorizontalAlignment = xlHAlignCenter
                                                                                       
   End With
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub




Private Sub Form_Load()

Dim r_str_CadAux As String
   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   ipp_FecIni.Text = "01/01/2016" 'modctb_str_FecIni
   ipp_FecFin.Text = Format(Now, "dd/MM/yyyy")
End Sub
