Attribute VB_Name = "ImportaCheques"
Option Explicit


Public Function FCod_DIA()
' Calcula o digito do dia do movimento do sistema

Dim F_DIA1, F_DIA2, F_MES1, F_MES2, F_ANO1, F_ANO2 As Integer
Dim F_DIGAUX1, F_DIGAUX2, F_TOTDIG1, F_DIGITO1, F_DIGITO2, F_TOTDIG2 As Integer
Dim cData As String

 cData = CStr(Geral.DataProcessamento)


 F_DIA1 = CInt(Mid(cData, 7, 1))
 F_DIA2 = CInt(Mid(cData, 8, 1))

 F_MES1 = CInt(Mid(cData, 5, 1))
 F_MES2 = CInt(Mid(cData, 6, 1))
 
 F_ANO1 = CInt(Mid(cData, 3, 1))
 F_ANO2 = CInt(Mid(cData, 4, 1))
 
 
 F_TOTDIG1 = IIf(F_DIA1 * 1 > 9, (F_DIA1 * 1) - 9, (F_DIA1 * 1)) + _
            IIf(F_DIA2 * 2 > 9, (F_DIA2 * 2) - 9, (F_DIA2 * 2)) + _
            IIf(F_MES1 * 1 > 9, (F_MES1 * 1) - 9, (F_MES1 * 1)) + _
            IIf(F_MES2 * 2 > 9, (F_MES2 * 2) - 9, (F_MES2 * 2)) + _
            IIf(F_ANO1 * 1 > 9, (F_ANO1 * 1) - 9, (F_ANO1 * 1)) + _
            IIf(F_ANO2 * 2 > 9, (F_ANO2 * 2) - 9, (F_ANO2 * 2))

F_DIGAUX1 = CInt(Mid(Format(CStr(F_TOTDIG1), "00"), 2, 1))

F_DIGAUX1 = 10 - F_DIGAUX1
F_DIGITO1 = CStr(IIf(F_DIGAUX1 > 9, 0, F_DIGAUX1))
  
F_TOTDIG2 = IIf(F_DIA1 * 2 > 9, (F_DIA1 * 2) - 9, (F_DIA1 * 2)) + _
            IIf(F_DIA2 * 1 > 9, (F_DIA2 * 1) - 9, (F_DIA2 * 1)) + _
            IIf(F_MES1 * 2 > 9, (F_MES1 * 2) - 9, (F_MES1 * 2)) + _
            IIf(F_MES2 * 1 > 9, (F_MES2 * 1) - 9, (F_MES2 * 1)) + _
            IIf(F_ANO1 * 2 > 9, (F_ANO1 * 2) - 9, (F_ANO1 * 2)) + _
            IIf(F_ANO2 * 1 > 9, (F_ANO2 * 1) - 9, (F_ANO2 * 1))
  
F_DIGAUX2 = CInt(Mid(Format(CStr(F_TOTDIG2), "00"), 2, 1))
F_DIGAUX2 = 10 - F_DIGAUX2
  
F_DIGITO2 = CStr(IIf(F_DIGAUX2 > 9, 0, F_DIGAUX2))

FCod_DIA = (F_DIGITO1 + F_DIGITO2)

End Function

Public Function F_Inidk(nNumRem As Integer)

Dim cData As String
Dim cCreTemp As String
Dim Total As String
Dim CodAplic As Integer

cData = CStr(Geral.DataProcessamento)
      
cData = Right(cData, 2) + Mid(cData, 5, 2)

CodAplic = 0
Total = Format(CStr(CInt(CodAplic) + nNumRem + CInt(cData)), "000000000000000")


cCreTemp = Space(3) & _
          g_Parametros.CodigoAplicacao & _
          Space(1) & _
          Format(CStr(nNumRem), "0000") & _
          Space(1) & _
          cData & _
          Space(1) & _
          Total & _
          Space(11) & _
          Space(32) & _
          "INIDK" & _
          Chr(13) + Chr(10)
 
F_Inidk = cCreTemp
End Function

Public Function F_Inik7(nNumRem As Integer)
Dim cData, cCreTemp, cCodDia, Total As String


cData = CStr(Geral.DataProcessamento)
       
cData = Right(cData, 2) + Mid(cData, 5, 2)

cCodDia = FCod_DIA()

Total = Format(CStr(CInt(cCodDia) + 0 + nNumRem + CInt(cData)), "0000")

cCreTemp = cCodDia & _
          Space(1) & _
          g_Parametros.CodigoAplicacao & _
          Space(1) & _
          Format(CStr(nNumRem), "0000") & _
          Space(1) & _
          cData & _
          Space(1) & _
          Format(CStr(g_Parametros.Codigo_USB), "0000") & _
          "0001" & _
          "001" & _
          Total & _
          Space(11) & _
          Format(g_Parametros.CPD_Origem, "0000") & _
          Space(27) & _
          "0INIK7" & _
          Chr(13) + Chr(10)
          
                    
F_Inik7 = cCreTemp
End Function


Public Function F_HDX(nNumRem As Integer)
Dim cData, cCreTemp, Total As String

cData = Right(CStr(Geral.DataProcessamento), 6)

cCreTemp = Format(CStr(nNumRem), "0000") & _
          "FC" & _
          "HDX0" & _
          Format(g_Parametros.CPD_Origem, "0000") & _
          g_Parametros.CodigoAplicacao & _
          Format(g_Parametros.CPD_Destino, "0000") & _
          g_Parametros.CodigoAplicacao & _
          "1000" & _
          g_Parametros.Codigo_Terceira + Format(CStr(nNumRem), "0000") + "GCC" & _
          cData & _
          "0000000" & _
          Chr(13) + Chr(10)
F_HDX = cCreTemp
End Function

