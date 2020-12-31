Attribute VB_Name = "AvisoDiferenca"
Option Explicit
' Registro de Aviso de Diferença -  Rótulo CHADIF
Public Type AvisoDif_Reg
    CodigoOcorrencia                As String * 9
    DataOcorrencia                  As String * 8
    DataDeposito                    As String * 8
    Num_Bordero                     As String * 18
    CodigoCarteira                  As String * 2
    Agencia                         As String * 4
    Conta                           As String * 7
    CodigoDevolucao                 As String * 2
    CodigoCompensacao               As String * 3
    BancoEmitente                   As String * 4
    AgenciaEmitente                 As String * 4
    CcEmitente                      As String * 11
    NrChequeEmitente                As String * 10
    TipoCheque                      As String * 1
    TipoInscricao                   As String * 2
    InscricaoEmitente               As String * 14
    Valor                           As String * 13
    MotivoDevolucao                 As String * 20
    Gerado                          As String * 1
    CrLf                            As String * 2
End Type


Public Function Ler_AvisoDif(ByVal PathName As String) As Boolean
 
    Dim rstAvisoDif         As New ADODB.Recordset
    Dim InsAviso            As New Custodia.Inserir
    Dim SelAviso            As New Custodia.Selecionar
    Dim DatFile             As Integer
    Dim vMotivo             As Integer
    Dim vPos                As Integer
    Dim sCodOcorrencia      As String
    Dim lRetorno            As Long
    Dim Reg                 As String * 196
    Dim OffSet              As Long
    Dim nAvisos             As Long
    Dim AD                  As AvisoDif_Reg
    Dim sStr                As String
    Dim sWhere              As String
    Dim Progress            As New clsProgressBar
    Dim lngRegs             As Long
    
    On Error GoTo ErroLeitura
        
    Ler_AvisoDif = False
    nAvisos = 0
    vPos = 1
    
    Screen.MousePointer = vbHourglass
    
    If Not FileExist(Trim(g_Parametros.DiretorioRecepcao) & PathName) Then
         MsgBox "Arquivo de Aviso de Direfença Não Encontrado", vbOKOnly + vbExclamation, App.Title
         GoTo FimLeitura
    End If
    
    DatFile = FreeFile
    Open Trim(g_Parametros.DiretorioRecepcao) & PathName For Binary Access Read Lock Read Write As #DatFile
        
    OffSet = 1
    
    Get #DatFile, OffSet, Reg
    
    'Obtem o total de registros do arquivo de leitura
    If Not EOF(DatFile) Then
          'Inicia progress bar
          Progress.ValorMinimo = 1
          Progress.ValorMaximo = Fix(FileLen(Trim(g_Parametros.DiretorioRecepcao) & PathName) / Len(Reg))
          Progress.DescricaoProcesso = "Recepcionando Aviso de Diferença ..."
          Progress.InicializaProgressBar
          lngRegs = 0
    End If
          
    While Not EOF(DatFile)
        'Acumulador de registros lidos
        lngRegs = lngRegs + 1
        
        ' Se arquivo foi lido ok
        If Len(Reg) < 194 Then
             MsgBox "Erro de Leitura", vbOKOnly + vbCritical, App.Title
             GoTo FimLeitura
        End If
        
        ' Ver rótulo do arquivo
        If Mid(Reg, 1, 6) <> "CHADIF" Then
             MsgBox "Rótulo do Arquivo de Diferença Inválido.", vbOKOnly + vbCritical, App.Title
             GoTo FimLeitura
        End If
        
        ' Ver se CGC de terceira é válido
        If CStr(Mid(Reg, 7, 14)) <> g_Parametros.CNPJ_Terceira Then
             MsgBox "CNPJ da Terceira Inválido", vbOKOnly + vbCritical, App.Title
             GoTo FimLeitura
        End If
                
        ' Atribuir registos do arquivo
        AD.DataOcorrencia = Mid(Reg, 41, 8)
        AD.CodigoOcorrencia = Mid(Reg, 49, 9)
        AD.MotivoDevolucao = Mid(Reg, 175, 20)
        AD.DataDeposito = Mid(Reg, 120, 8)
        AD.Num_Bordero = Mid(Reg, 21, 18)
        AD.CodigoCarteira = Mid(Reg, 39, 2)
        AD.Agencia = Mid(Reg, 58, 4)
        AD.Conta = Mid(Reg, 62, 7)
        AD.CodigoDevolucao = Mid(Reg, 69, 2)
        AD.CodigoCompensacao = Mid(Reg, 71, 3)
        AD.BancoEmitente = Mid(Reg, 74, 4)
        AD.AgenciaEmitente = Mid(Reg, 78, 4)
        AD.CcEmitente = Mid(Reg, 82, 11)
        AD.NrChequeEmitente = Mid(Reg, 93, 10)
        AD.TipoCheque = Mid(Reg, 103, 1)
        AD.TipoInscricao = Mid(Reg, 104, 2)
        AD.InscricaoEmitente = Mid(Reg, 106, 14)
        AD.Valor = Mid(Reg, 128, 13)
        AD.Gerado = 0
        
        
        For vMotivo = 1 To 10
        
        sCodOcorrencia = Mid(AD.MotivoDevolucao, vPos, 2)
        
        If sCodOcorrencia = Space(2) Then
          Exit For
        End If
        
        Set rstAvisoDif = g_cMainConnection.Execute(SelAviso.GetAvisoDif(CLng(AD.DataOcorrencia), CInt(AD.CodigoOcorrencia), CInt(sCodOcorrencia)))
        
        If rstAvisoDif.EOF Then
        
          ' Gravar Registro do Aviso de Diferença
          
          Call g_cMainConnection.Execute(InsAviso.InsereAvisoDiferenca(CLng(AD.DataOcorrencia), _
                                         CLng(AD.CodigoOcorrencia), CInt(sCodOcorrencia), CLng(AD.DataDeposito), _
                                         AD.Num_Bordero, CByte(AD.CodigoCarteira), _
                                         CInt(AD.Agencia), CLng(AD.Conta), CInt(AD.CodigoDevolucao), _
                                         CInt(AD.CodigoCompensacao), CInt(AD.BancoEmitente), _
                                         CInt(AD.AgenciaEmitente), CLng(AD.CcEmitente), _
                                         CLng(AD.NrChequeEmitente), CByte(AD.TipoCheque), _
                                         CByte(AD.TipoInscricao), AD.InscricaoEmitente, _
                                         Format(Val(InserePonto(AD.Valor)), MASK_VALOR), AD.Gerado), _
                                         lRetorno, adCmdText)
                   
            nAvisos = nAvisos + lRetorno
            vPos = vPos + 2
            
        End If
        
        
        Next
        
        OffSet = OffSet + Len(Reg)
        
        Get #DatFile, OffSet, Reg
        vPos = 1
        
        'Atualiza Progress Bar
        Progress.AtualValue = lngRegs
        Progress.AtualizaBarra
        
    Wend
    
    Close #DatFile
    Ler_AvisoDif = True
    
    'Encerra progress bar
    Set Progress = Nothing
    
    Screen.MousePointer = vbDefault
    MsgBox "Foram Processados " & CStr(nAvisos) & " Avisos de Diferença.", vbOKOnly + vbExclamation, App.Title
    
    Exit Function
    
FimLeitura:

    Close #DatFile
    'Encerra progress bar
    Set Progress = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Function
    
ErroLeitura:
    
    MsgBox "Erro na Leitura do Arquivo de Aviso de Diferença.", vbOKOnly + vbCritical, App.Title
    
    'Encerra progress bar
    Set Progress = Nothing
    
    Screen.MousePointer = vbDefault
End Function
