Attribute VB_Name = "TiposDados"
Option Explicit

Public Enum enumVipsDll
    eDllUnibanco = 0
    eDllProservi
End Enum

''''''''''''''''''''''''''''''''''''''''''
' Versão 3.3 (66)                        '
' Escolher o tipo de autenticadora usada '
''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''
' Tipos de Autenticadora '
''''''''''''''''''''''''''
Public Enum enumAutentica
    estAutentica = "0"
    estAutIBM = "1"
    estAutProcomp = "2"
End Enum

''''''''''''''''''''
' Tipos de Scanner '
''''''''''''''''''''
Public Enum enumScanner
    escnDummy = -1
    escnSemScanner
    escnVIPS
    escnCanonLS500
    escnLS500
    escnCanon
End Enum
''''''''''''''''''''''
' Status do envelope '
''''''''''''''''''''''
Public Enum enumStatucEnvelope
    estenvCadastrado = "0"
    estenvDigitar = "1"
    estenvDigitando = "2"
End Enum
'''''''''''''''''''''''''''''
' Public Tipos de Documento '
'''''''''''''''''''''''''''''
Public Enum enumTipoDocto
    etpdocDesconhecido = 0
    etpdocEnvelope                      '01
    etpdocDepositoCC                    '02
    etpdocDepositoCP                    '03
    etpdocADCC                          '04
    etpdocChequeUBBSacado               '05
    etpdocChequeTerceiroPagto           '06
    etpdocChequeDeposito                '07
    etpdocConcessionariaValorReais      '08
    etpdocConcessionariaValorIndexado   '09
    etpdocFichaCompensacao              '10
    etpdocInss                          '11     (Retirado)
    etpdocTitulos                       '12
    etpdocCobRegistrada                 '13
    etpdocCobEspecial                   '14
    etpdocDarm                          '15
    etpdocDarfPreto                     '16
    etpdocDarfSimples                   '17
    etpdocGare                          '18
    etpdocGRPS                          '19     (Retirado)
    etpdocAgua                          '20
    etpdocGas                           '21
    etpdocLuz                           '22
    etpdocTelefone                      '23
    etpdocTributosMunicipais            '24
    etpdocTributosEstaduais             '25
    etpdocTributosFederais              '26
    etpdocArrecConvencional             '27
    etpdocUnicobrancaUBB                '28
    etpdocCobrancaImediataUBB           '29
    etpdocCobrancaEspecialUBB           '30
    etpdocCobrancaTerceiros             '31
    etpdocAcCredDep                     '32
    etpdocAcDebDep                      '33
    etpdocAcCredCh                      '34
    etpdocGPS                           '35
    etpdocCartaoAvulso                  '36
    etpdocOCT                           '37
    etpdocDebAuto                       '38
    etpdocCapaOCT                       '39
    etpdocFGTS                          '40
    etpdocLancamentoInterno             '41
    etpdocMalote = 99                   '99

End Enum

