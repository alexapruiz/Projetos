Attribute VB_Name = "Tipos_ControleGeracao"
Option Explicit

''''''''''''''''''''''''''
'Nome dos arquivos de I/O'
''''''''''''''''''''''''''
Public Const ARQ_DADOS = "DADOS.DAT"
Public Const ARQ_AGENCIA = "AGENCIA.DAT"
Public Const ARQ_CD = "CD.ID"

''''''''''''''''''''''''''''''''
'Nome dos diretorios de leitura'
''''''''''''''''''''''''''''''''
Public Const DIR_DADOS = "DADOS\"
Public Const DIR_IMAGENS = "IMAGENS\"


''''''''''''''''''''''''''''''''''''''''''''
'Nome do diretorio onde contera os arquivos'
'sempre abaixo de \mdi_ag\dados\
''''''''''''''''''''''''''''''''''''''''''''
Public Const DIR_CD = "CD\"



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                             campos comuns                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Campos do Header
Public Type cg_Header
    TipoRegistro                    As String * 1
    Sequencial                      As String * 6
    DataProcessamento               As String * 8
    AgOrig                          As String * 4
    Remessa                         As String * 5
    CrLf                            As String * 2
End Type
'Campos do Trailler
Public Type cg_Trailler
    TipoRegistro                    As String * 1
    Sequencial                      As String * 6
    DataProcessamento               As String * 8
    AgOrig                          As String * 4
    Remessa                         As String * 5
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                  Campos do arquivo Agencia.Dat                                            '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Campos do registro da agencia
Public Type cg_Registro
    TipoRegistro                    As String * 1
    Sequencial                      As String * 6
    Agencia                         As String * 4
    Lacre                           As String * 8
    QtdInformada                    As String * 6
    HoraCadastrada                  As String * 5
    IdEnv_Mal                       As String * 1
    CrLf                            As String * 2
End Type

Public Type cg_AGENCIA
    Header                          As cg_Header
    Registro                        As cg_Registro
    Trailler                        As cg_Trailler
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                   Estrutura do arquivo Dados.Dat                                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Campos do Lote
Public Type cg_Lote
    TipoRegistro                    As String * 1
    Sequencial                      As String * 6
    IdLote                          As String * 9
    Status                          As String * 1
    Prioridade                      As String * 1
    CrLf                            As String * 2
End Type

'Campos do Log
Public Type cg_Log
    TipoRegistro                    As String * 1
    Sequencial                      As String * 6
    Data                            As String * 17
    Login                           As String * 10
    Acao                            As String * 3
    CrLf                            As String * 2
End Type

'Campos da capa
Public Type cg_Capa
    TipoRegistro                    As String * 1
    Sequencial                      As String * 6
    IdLote                          As String * 9
    IdEnv_Mal                       As String * 1
    Capa                            As String * 14
    Num_Malote                      As String * 9
    AgOrig                          As String * 4
    Status                          As String * 1
    Ocorrencia                      As String * 3
    Duplicidade                     As String * 1
    CrLf                            As String * 2
End Type

'Campos do documento
Public Type cg_Documento
    TipoRegistro                    As String * 1
    Sequencial                      As String * 6
    OrdemCaptura                    As String * 5
    TipoDocto                       As String * 2
    Leitura                         As String * 48
    Frente                          As String * 20
    Verso                           As String * 20
    Status                          As String * 1
    Ordem                           As String * 1
    CrLf                            As String * 2
End Type


'Todos os registros
Public Type cg_DADOS
    Header                          As cg_Header
    Lote                            As cg_Lote
    Capa                            As cg_Capa
    Docto                           As cg_Documento
    Trailler                        As cg_Trailler
    Log                             As cg_Log
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                   Estrutura do arquivo CD.ID                                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Type cg_CD
    Agencia                         As String * 4
    Data                            As String * 8
    Hora                            As String * 8 'HH:MM:SS
    Remessa                         As String * 5
    Numero_CD                       As String * 2
End Type
