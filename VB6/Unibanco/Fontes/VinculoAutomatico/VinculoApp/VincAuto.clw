; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CGetDocOcorrencia
LastTemplate=CRecordset
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "VincAuto.h"

ClassCount=8
Class1=CVincAutoApp
Class2=CVincAutoDlg

ResourceCount=3
Resource2=IDD_VINCAUTO_DIALOG
Resource1=IDR_MAINFRAME
Class3=CGetDocumentos
Class4=CGetAgContaDeposito
Class5=CGetIdDoctoAjuste
Class6=CGetDocOcorrencia
Class7=CGetDocumentoTransmitido
Class8=CGetControleCapa
Resource3=IDD_VINCAUTO_DIALOG (English (U.S.))

[CLS:CVincAutoApp]
Type=0
HeaderFile=VincAuto.h
ImplementationFile=VincAuto.cpp
Filter=N
LastObject=CVincAutoApp

[CLS:CVincAutoDlg]
Type=0
HeaderFile=VincAutoDlg.h
ImplementationFile=VincAutoDlg.cpp
Filter=D
BaseClass=CDialog
VirtualFilter=dWC
LastObject=IDC_BUTTON_DONE



[DLG:IDD_VINCAUTO_DIALOG]
Type=1
Class=CVincAutoDlg
ControlCount=5
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_BUTTON_INIT,button,1342242816
Control4=IDC_BUTTON_DONE,button,1342242816
Control5=IDC_BUTTON_EXEC,button,1342242816

[DLG:IDD_VINCAUTO_DIALOG (English (U.S.))]
Type=1
Class=CVincAutoDlg
ControlCount=5
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_BUTTON_INIT,button,1342242816
Control4=IDC_BUTTON_DONE,button,1342242816
Control5=IDC_BUTTON_EXEC,button,1342242816

[CLS:CGetDocumentos]
Type=0
HeaderFile=GetDocumentos.h
ImplementationFile=GetDocumentos.cpp
BaseClass=CRecordset
Filter=N
VirtualFilter=r
LastObject=CGetDocumentos

[DB:CGetDocumentos]
DB=1
DBType=ODBC
ColumnCount=25
Column1=[DataProcessamento], 4, 4
Column2=[IdDocto], 4, 4
Column3=[IdCapa], 4, 4
Column4=[TipoDocto], 5, 2
Column5=[Ocorrencia], 3, 7
Column6=[Leitura], 12, 48
Column7=[Frente], 12, 20
Column8=[Verso], 12, 20
Column9=[Status], 1, 1
Column10=[Alcada], 1, 1
Column11=[Autenticado], 1, 2
Column12=[Cortado], 1, 1
Column13=[OcorrenciaOk], 1, 1
Column14=[Ordem], 1, 1
Column15=[NSU], 1, 6
Column16=[Terminal], 5, 2
Column17=[Vinculo], 4, 4
Column18=[AgenciaVinculo], 4, 4
Column19=[ContaVinculo], 2, 9
Column20=[CMC7Associado], 1, 30
Column21=[Duplicidade], 2, 3
Column22=[CodCenape], 2, 11
Column23=[CodBarComplem], 1, 1
Column24=[Valor], 3, 21
Column25=[Atualizacao], -2, 8

[CLS:CGetAgContaDeposito]
Type=0
HeaderFile=GetAgContaDeposito.h
ImplementationFile=GetAgContaDeposito.cpp
BaseClass=CRecordset
Filter=N
VirtualFilter=r
LastObject=CGetAgContaDeposito

[DB:CGetAgContaDeposito]
DB=1
DBType=ODBC
ColumnCount=11
Column1=[DataProcessamento], 4, 4
Column2=[IdDocto], 4, 4
Column3=[CMC7], 12, 30
Column4=[Identificado], 2, 8
Column5=[Agencia], 5, 2
Column6=[Conta], 2, 9
Column7=[TipoConta], -6, 1
Column8=[Dinheiro], 3, 21
Column9=[Cheque], 3, 21
Column10=[Valor], 3, 21
Column11=[Atualizacao], -2, 8

[DB:ObtemIdDoctoAjuste]
DB=1
DBType=ODBC
ColumnCount=25
Column1=[DataProcessamento], 4, 4
Column2=[IdDocto], 4, 4
Column3=[IdCapa], 4, 4
Column4=[TipoDocto], 5, 2
Column5=[Ocorrencia], 3, 7
Column6=[Leitura], 12, 48
Column7=[Frente], 12, 20
Column8=[Verso], 12, 20
Column9=[Status], 1, 1
Column10=[Alcada], 1, 1
Column11=[Autenticado], 1, 2
Column12=[Cortado], 1, 1
Column13=[OcorrenciaOk], 1, 1
Column14=[Ordem], 1, 1
Column15=[NSU], 1, 6
Column16=[Terminal], 5, 2
Column17=[Vinculo], 4, 4
Column18=[AgenciaVinculo], 4, 4
Column19=[ContaVinculo], 2, 9
Column20=[CMC7Associado], 1, 30
Column21=[Duplicidade], 2, 3
Column22=[CodCenape], 2, 11
Column23=[CodBarComplem], 1, 1
Column24=[Valor], 3, 21
Column25=[Atualizacao], -2, 8

[CLS:CGetIdDoctoAjuste]
Type=0
HeaderFile=GetIdDoctoAjuste.h
ImplementationFile=GetIdDoctoAjuste.cpp
BaseClass=CRecordset
Filter=N
VirtualFilter=r
LastObject=CGetIdDoctoAjuste

[DB:CGetIdDoctoAjuste]
DB=1
DBType=ODBC
ColumnCount=25
Column1=[DataProcessamento], 4, 4
Column2=[IdDocto], 4, 4
Column3=[IdCapa], 4, 4
Column4=[TipoDocto], 5, 2
Column5=[Ocorrencia], 3, 7
Column6=[Leitura], 12, 48
Column7=[Frente], 12, 20
Column8=[Verso], 12, 20
Column9=[Status], 1, 1
Column10=[Alcada], 1, 1
Column11=[Autenticado], 1, 2
Column12=[Cortado], 1, 1
Column13=[OcorrenciaOk], 1, 1
Column14=[Ordem], 1, 1
Column15=[NSU], 1, 6
Column16=[Terminal], 5, 2
Column17=[Vinculo], 4, 4
Column18=[AgenciaVinculo], 4, 4
Column19=[ContaVinculo], 2, 9
Column20=[CMC7Associado], 1, 30
Column21=[Duplicidade], 2, 3
Column22=[CodCenape], 2, 11
Column23=[CodBarComplem], 1, 1
Column24=[Valor], 3, 21
Column25=[Atualizacao], -2, 8

[CLS:CGetDocOcorrencia]
Type=0
HeaderFile=GetDocOcorrencia.h
ImplementationFile=GetDocOcorrencia.cpp
BaseClass=CRecordset
Filter=N
VirtualFilter=r
LastObject=CGetDocOcorrencia

[DB:CGetDocOcorrencia]
DB=1
DBType=ODBC
ColumnCount=24
Column1=[DataProcessamento], 4, 4
Column2=[IdDocto], 4, 4
Column3=[IdCapa], 4, 4
Column4=[TipoDocto], 5, 2
Column5=[Ocorrencia], 3, 7
Column6=[Leitura], 12, 48
Column7=[Frente], 12, 20
Column8=[Verso], 12, 20
Column9=[Status], 1, 1
Column10=[Alcada], 1, 1
Column11=[Autenticado], 1, 2
Column12=[Cortado], 1, 1
Column13=[OcorrenciaOk], 1, 1
Column14=[Ordem], 1, 1
Column15=[NSU], 1, 6
Column16=[Terminal], 5, 2
Column17=[Vinculo], 4, 4
Column18=[AgenciaVinculo], 4, 4
Column19=[ContaVinculo], 2, 9
Column20=[CMC7Associado], 1, 30
Column21=[Duplicidade], 2, 3
Column22=[CodCenape], 2, 11
Column23=[CodBarComplem], 1, 1
Column24=[Valor], 3, 21

[CLS:CGetDocumentoTransmitido]
Type=0
HeaderFile=GetDocumentoTransmitido.h
ImplementationFile=GetDocumentoTransmitido.cpp
BaseClass=CRecordset
Filter=N
LastObject=CGetDocumentoTransmitido

[DB:CGetDocumentoTransmitido]
DB=1
DBType=ODBC
ColumnCount=24
Column1=[DataProcessamento], 4, 4
Column2=[IdDocto], 4, 4
Column3=[IdCapa], 4, 4
Column4=[TipoDocto], 5, 2
Column5=[Ocorrencia], 3, 7
Column6=[Leitura], 12, 48
Column7=[Frente], 12, 20
Column8=[Verso], 12, 20
Column9=[Status], 1, 1
Column10=[Alcada], 1, 1
Column11=[Autenticado], 1, 2
Column12=[Cortado], 1, 1
Column13=[OcorrenciaOk], 1, 1
Column14=[Ordem], 1, 1
Column15=[NSU], 1, 6
Column16=[Terminal], 5, 2
Column17=[Vinculo], 4, 4
Column18=[AgenciaVinculo], 4, 4
Column19=[ContaVinculo], 2, 9
Column20=[CMC7Associado], 1, 30
Column21=[Duplicidade], 2, 3
Column22=[CodCenape], 2, 11
Column23=[CodBarComplem], 1, 1
Column24=[Valor], 3, 21

[DB:CGetCapaCSP]
DB=1
DBType=ODBC
ColumnCount=0

[CLS:CGetControleCapa]
Type=0
HeaderFile=GetControleCapa.h
ImplementationFile=GetControleCapa.cpp
BaseClass=CRecordset
Filter=N
VirtualFilter=r

[DB:CGetControleCapa]
DB=1
DBType=ODBC
ColumnCount=21
Column1=[DataProcessamento], 4, 4
Column2=[IdCapa], 4, 4
Column3=[IdLote], 4, 4
Column4=[idEnv_Mal], 1, 1
Column5=[Capa], 2, 20
Column6=[Num_Malote], 2, 13
Column7=[AgOrig], 5, 2
Column8=[Status], 1, 1
Column9=[DataCriacao], 11, 16
Column10=[Alcada], 1, 1
Column11=[PendenciaValor], 1, 1
Column12=[Supervisor], 1, 1
Column13=[VinculoManual], 1, 1
Column14=[IgnorarProva0], 1, 1
Column15=[Ocorrencia], 3, 7
Column16=[Conta], 3, 21
Column17=[Dinheiro], 3, 21
Column18=[Diferenca], 3, 21
Column19=[Duplicidade], 2, 3
Column20=[HoraAtual], 11, 16
Column21=[RecepcionadoIK], 1, 1

