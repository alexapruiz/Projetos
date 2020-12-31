Attribute VB_Name = "basCripto"
Option Explicit
Public Key(3) As Integer
Private Declare Function Encripta Lib "Encripta.dll" (ByVal lngIn As Long, ByVal strOut As String) As Long
Public Function Encript(ByVal aString As String) As String
    Dim Result(0 To 255) As Byte
    Dim Count As Long
    Dim Remainder As Integer
    Dim Quocient As Integer
    Dim Retorno As String
    
    Quocient = Len(aString) \ 4
    Remainder = Len(aString) Mod 4
    For Count = 0 To Quocient - 1
        Result(Count * 4 + 0) = (Asc(Mid(aString, Count * 4 + 1, 1)) + Key(0)) Mod 255
        Result(Count * 4 + 1) = (Asc(Mid(aString, Count * 4 + 2, 1)) + Key(1)) Mod 255
        Result(Count * 4 + 2) = (Asc(Mid(aString, Count * 4 + 3, 1)) + Key(2)) Mod 255
        Result(Count * 4 + 3) = (Asc(Mid(aString, Count * 4 + 4, 1)) + Key(3)) Mod 255
        
        Result(Count * 4 + 0) = Result(Count * 4 + 0) Xor Key(0)
        Result(Count * 4 + 1) = Result(Count * 4 + 1) Xor Key(1)
        Result(Count * 4 + 2) = Result(Count * 4 + 2) Xor Key(2)
        Result(Count * 4 + 3) = Result(Count * 4 + 3) Xor Key(3)
    Next
    If Remainder > 0 Then
        For Count = 0 To Remainder - 1
            Result(Quocient * 4 + Count) = (Asc(Mid(aString, Quocient * 4 + Count + 1, 1)) + Key(Count)) Mod 255
            Result((Quocient * 4) + Count) = Result(Quocient * 4 + Count) Xor Key(Count)
        Next
    End If
    For Count = 0 To Len(aString) - 1
        Retorno = Retorno + Chr(Result(Count))
    Next
    Encript = Retorno
End Function
Public Function Decript(ByVal aString As String) As String
    Dim Result(0 To 255) As Byte
    Dim Count As Long
    Dim Remainder As Integer
    Dim Quocient As Integer
    Dim Retorno As String
    
    Quocient = Len(aString) \ 4
    Remainder = Len(aString) Mod 4
    For Count = 0 To Quocient - 1
        Result(Count * 4 + 0) = Asc(Mid(aString, Count * 4 + 1, 1)) Xor Key(0)
        Result(Count * 4 + 1) = Asc(Mid(aString, Count * 4 + 2, 1)) Xor Key(1)
        Result(Count * 4 + 2) = Asc(Mid(aString, Count * 4 + 3, 1)) Xor Key(2)
        Result(Count * 4 + 3) = Asc(Mid(aString, Count * 4 + 4, 1)) Xor Key(3)
        
        Result(Count * 4 + 0) = (Result(Count * 4 + 0) - Key(0) + 255) Mod 255
        Result(Count * 4 + 1) = (Result(Count * 4 + 1) - Key(1) + 255) Mod 255
        Result(Count * 4 + 2) = (Result(Count * 4 + 2) - Key(2) + 255) Mod 255
        Result(Count * 4 + 3) = (Result(Count * 4 + 3) - Key(3) + 255) Mod 255
    Next
    If Remainder > 0 Then
        For Count = 0 To Remainder - 1
            Result(Quocient * 4 + Count) = Asc(Mid(aString, Quocient * 4 + Count + 1, 1)) Xor Key(Count)
            Result(Quocient * 4 + Count) = (Result(Quocient * 4 + Count) - Key(Count) + 255) Mod 255
        Next
    End If
    For Count = 0 To Len(aString) - 1
        Retorno = Retorno + Chr(Result(Count))
    Next
    Decript = Retorno
End Function
Public Function ValidaEncriptaBO() As Boolean
'-----------------------------------------------------------------------------------------------
'   Função:     EncriptaBO
'   DLL:        Encripta.DLL
'
'   Parâmetros: strTipodocto    - Tipo do documento
'               strCampoChave   - Informação do documento para transformação da chave de encriptamento
'               lngIdDocto      - (Opcional) IdDocto do documento, Se Informado Gravar Autenticação Digital
'
'   Retorno:    String encriptografada
'   Regra:      Segue abaixo relação do Documento, TipoDocto e Campos para formação da chave;
'               ADCC(4)                         TipoDocto + Conta
'               Ajuste Credito(32,34)           TipoDocto + Conta
'               Ajuste Debito(33,38)            TipoDocto + Conta
'               Arrec.Convencional(27)          TipoDocto + Produto
'               Arrec.Eletrônica(20,21,22,23)   TipoDocto + Subs(CdBarras,1,4) + Subs(CdBarras,11,3)
'               Cartão Avulso(36)               TipoDocto + Subs(Cartão,10,7)
'               CBIndex(24,25,26)               TipoDocto + Subs(CdBarras,1,4) + Subs(CdBarras,11,3)
'               Cheque(5,6,7)                   TipoDocto + Subs(CMC7,23,7)
'               Cobrança Especial(14)           TipoDocto + Cedente
'               Cobrança Registrada(13)         TipoDocto + Subs(NossoNumero,1,7)
'               DarfPreto(16)                   TipoDocto + Subs(CPFCGC,8,7)
'               DarfSimples(17)                 TipoDocto + Subs(CPFCGC,8,7)
'               Darm(15)                        TipoDocto + Subs(CCM,1,7)
'               Depósito(2,3)                   TipoDocto + Conta
'               FGTS(40)                        TipoDocto + Subs(CNPJ_Empresa,8,7)
'               Ficha Compensação(28,29,30,31)  TipoDocto + Subs(CdBarras,1,7)
'               GARE(18)                        ????????????
'               GPS(35)                         TipoDocto + Subs(Identificador,1,7)
'               Lançamento Interno(41)          TipoDocto + Subs(ControleBanco,1,7)
'               OCT(37)                         TipoDocto + ContaCredito
'               Título(12)                      TipoDocto + Banco
'-----------------------------------------------------------------------------------------------

On Error GoTo TrataErro
    
    Dim strEncripta     As String
    Dim strChave        As String
    Dim strTipodocto    As Byte
    Dim i               As Integer
    
    strEncripta = Space(16)

    strTipodocto = CStr(Geral.rstDoctos!TipoDocto)
    LocalLog "Rotina de Verificação de Chave de Encriptação Doc. " & strTipodocto
        
    Select Case strTipodocto
        Case 42, 43, 44, 45, 32, 33
            ValidaEncriptaBO = True
            Exit Function
        Case 2, 3, 4, 34, 38
            strChave = Right(String(8, "0") & CStr(strTipodocto) & CStr(Trim(Geral.rstDocto!Conta)), 8)
            
        Case 12
            strChave = Right(String(8, "0") & CStr(strTipodocto) & CStr(Trim(Geral.rstDocto!Banco)), 8)
            
        Case 14
            strChave = Right(String(8, "0") & CStr(strTipodocto) & CStr(Trim(Geral.rstDocto!Cedente)), 8)
            
        Case 27
            strChave = Right(String(8, "0") & CStr(strTipodocto) & CStr(Trim(Geral.rstDocto!Produto)), 8)
            
        Case 37
            strChave = Right(String(8, "0") & CStr(strTipodocto) & CStr(Trim(Geral.rstDocto!ContaCredito)), 8)
            
        Case 20, 21, 22, 23, 24, 25, 26
            strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDoctos!leitura, 1, 4) & Mid(Geral.rstDoctos!leitura, 11, 3), 8)
           'strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!CodigoBarras, 1, 4) & Mid(Geral.rstDocto!CodigoBarras, 11, 3), 8)
            
        Case 36
            strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!Cartao, 1, 7), 8)
            
        Case 5, 6, 7
            strChave = Right(String(8, "0") & CStr("5") & Mid(Geral.rstDoctos!leitura, 23, 7), 8)
        
        Case 13
            strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!NossoNumero, 1, 7), 8)
            
        Case 15
            strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!CCM, 1, 7), 8)
            
        Case 28, 29, 30, 31
            strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!CodigoBarras, 1, 7), 8)
            
        Case 35
            strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!Identificador, 1, 7), 8)
            
        Case 41
            strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Val(Right(Geral.rstDocto!ControleBanco, 14)), 1, 7), 8)
        
        Case 16
            strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!CPFCGC, 1, 7), 8)
            
        Case 17
            strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!CGC, 1, 7), 8)
            
        Case 18
            If Geral.rstDocto!CPFCGC = 0 Then
                strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!InscricaoEstadual, 1, 7), 8)
            Else
                strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!CPFCGC, 1, 7), 8)
            End If
        
        Case 40
            strChave = Right(String(8, "0") & CStr(strTipodocto) & Mid(Geral.rstDocto!CNPJCEI_Empresa, 1, 7), 8)
            
    End Select
      
    Encripta CLng(Val((strChave))), strEncripta
    LocalLog "Calculada: " & strEncripta & " Recebida: " & Geral.rstDoctos!AutenticacaoDigital
    
   'Validacao da chave recebida com a calculada
    If strEncripta = Geral.rstDoctos!AutenticacaoDigital Then
         ValidaEncriptaBO = True
    Else
        ValidaEncriptaBO = False
    End If

    Exit Function
    
TrataErro:
    LocalLog "Falha na ValidaEncriptaBO Erro: " & Err.Description
    
End Function

