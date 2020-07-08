Attribute VB_Name = "modctb"
Option Explicit

Public modctb_str_FecIni      As String
Public modctb_str_FecFin      As String
Public modctb_int_PerAno      As Integer
Public modctb_int_PerMes      As Integer
Public modctb_str_CodEmp      As String
Public modctb_str_NomEmp      As String
Public modctb_str_CodSuc      As String
Public modctb_str_NomSuc      As String
Public modctb_int_CodLib      As Integer
Public modctb_str_NomLib      As String
Public modctb_lng_NumAsi      As Long

Dim l_str_Formul     As String
Dim l_str_Caract     As String
Dim l_int_ConFor     As Integer

Public Function ff_Num_Calcul(ByVal p_Cadena As String) As Double
    l_str_Formul = p_Cadena
    l_int_ConFor = 0
    
    Call fs_Num_LeeCar
    
    ff_Num_Calcul = ff_Num_SumRes()
End Function

Public Sub fs_Num_LeeCar()
   Do
      l_int_ConFor = l_int_ConFor + 1
      
      If l_int_ConFor <= Len(l_str_Formul) Then
         l_str_Caract = Mid$(l_str_Formul, l_int_ConFor, 1)
      Else
         l_str_Caract = "Ñ"
      End If
      
      DoEvents
   Loop Until l_str_Caract <> " "
End Sub

Public Function ff_Num_SumRes() As Double
   Dim r_str_Operad As String
   Dim r_dbl_Result As Double
    
   r_dbl_Result = ff_Num_MulDiv()
   
   Do While l_str_Caract = "+" Or l_str_Caract = "-"
      r_str_Operad = l_str_Caract
      
      Call fs_Num_LeeCar
      
      If r_str_Operad = "+" Then
         'r_dbl_Result = r_dbl_Result + ff_Num_MulDiv()
         r_dbl_Result = r_dbl_Result + CDbl(Format(ff_Num_MulDiv(), "########0.000000"))
         r_dbl_Result = CDbl(Format(r_dbl_Result, "########0.000000"))
      End If
      
      If r_str_Operad = "-" Then
         'r_dbl_Result = r_dbl_Result - ff_Num_MulDiv()
         r_dbl_Result = r_dbl_Result - CDbl(Format(ff_Num_MulDiv(), "########0.000000"))
         
         r_dbl_Result = CDbl(Format(r_dbl_Result, "########0.000000"))
      End If
      
      DoEvents
   Loop
   
   ff_Num_SumRes = r_dbl_Result
End Function

Public Function ff_Num_MulDiv() As Double
   Dim r_str_Operad  As String
   Dim r_dbl_Result  As Double
   Dim r_dbl_Auxili  As Double
    
   r_dbl_Result = ff_Num_Negati()
   
   Do While l_str_Caract = "*" Or l_str_Caract = "/"
      r_str_Operad = l_str_Caract
      
      Call fs_Num_LeeCar
      
      If r_str_Operad = "*" Then
         r_dbl_Auxili = ff_Num_Negati()
         'r_dbl_Result = r_dbl_Result * ff_Num_Negati()
         r_dbl_Result = CDbl(Format(r_dbl_Result * r_dbl_Auxili, "#########0.000000"))
      End If
      
      If r_str_Operad = "/" Then
         r_dbl_Auxili = ff_Num_Negati()
         If r_dbl_Auxili > 0 Then
            'r_dbl_Result = r_dbl_Result / r_dbl_Auxili
            
            r_dbl_Result = CDbl(Format(r_dbl_Result / r_dbl_Auxili, "#########0.000000"))
         End If
      End If
      
      DoEvents
   Loop
   
   ff_Num_MulDiv = r_dbl_Result
End Function

Public Function ff_Num_Negati() As Double
   If l_str_Caract = "-" Then
      ff_Num_Negati = -1 * ff_Num_Operac()
   Else
      ff_Num_Negati = ff_Num_Operac()
   End If
End Function

Public Function ff_Num_Operac() As Double
   Dim r_int_Inicio  As Integer
   Dim r_dbl_Result  As Double
    
   If (l_str_Caract >= "0" And l_str_Caract <= "9") Or l_str_Caract = "." Then
      r_int_Inicio = l_int_ConFor
      
      Do
         Call fs_Num_LeeCar
         
         DoEvents
      Loop Until Not ((l_str_Caract >= "0" And l_str_Caract <= "9") Or l_str_Caract = ".")
      
      If l_str_Caract = "." Then
         Do
            Call fs_Num_LeeCar
            
            DoEvents
         Loop Until Not ((l_str_Caract >= "0" And l_str_Caract <= "9") Or l_str_Caract = ".")
      End If
      
      If l_str_Caract = "E" Then
         Do
            Call fs_Num_LeeCar
            
            DoEvents
         Loop Until Not ((l_str_Caract >= "0" And l_str_Caract <= "9") Or l_str_Caract = ".")
      End If
      
      
      r_dbl_Result = CDbl(Mid$(l_str_Formul, r_int_Inicio, l_int_ConFor - r_int_Inicio))
   Else
      If l_str_Caract = "(" Then
         Call fs_Num_LeeCar
         
         r_dbl_Result = ff_Num_SumRes()
         
         If l_str_Caract = ")" Then
            Call fs_Num_LeeCar
         End If
      End If
   End If
   
   ff_Num_Operac = r_dbl_Result
End Function

Public Function modctb_gf_Genera_NumAsi(ByVal p_CodEmp As String, ByVal p_CodSuc As String, ByVal p_PerAno As Integer, ByVal p_PerMes As Integer, ByVal p_NumLib As Integer) As Long
   Dim r_lng_NumMov     As Long
   
   modctb_gf_Genera_NumAsi = 0
   
   g_str_Parame = "SELECT * FROM CTB_FOLCOM WHERE "
   g_str_Parame = g_str_Parame & "FOLCOM_CODEMP = '" & p_CodEmp & "' AND "
   g_str_Parame = g_str_Parame & "FOLCOM_CODSUC = '" & p_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "FOLCOM_PERANO = " & CStr(p_PerAno) & " AND "
   g_str_Parame = g_str_Parame & "FOLCOM_PERMES = " & CStr(p_PerMes) & " AND "
   g_str_Parame = g_str_Parame & "FOLCOM_CODLIB = " & CStr(p_NumLib) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      r_lng_NumMov = 1
   Else
      r_lng_NumMov = g_rst_Genera!FOLCOM_NUMERO + 1
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CTB_FOLCOM ("
      g_str_Parame = g_str_Parame & "'" & p_CodEmp & "', "
      g_str_Parame = g_str_Parame & "'" & p_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(p_PerAno) & ", "
      g_str_Parame = g_str_Parame & CStr(p_PerMes) & ", "
      g_str_Parame = g_str_Parame & CStr(p_NumLib) & ", "
      g_str_Parame = g_str_Parame & CStr(r_lng_NumMov) & ", "
      
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If r_lng_NumMov = 1 Then
         g_str_Parame = g_str_Parame & "1) "
      Else
         g_str_Parame = g_str_Parame & "2) "
      End If
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CTB_FOLCOM. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   modctb_gf_Genera_NumAsi = r_lng_NumMov
End Function


