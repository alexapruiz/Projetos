; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CGetControleCapa
LastTemplate=CRecordset
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "vinculomsg.h"
LastPage=0

ClassCount=4

ResourceCount=0
Class1=CGetDataProc
Class2=CGetValoresParametro
Class3=CGetCapaVincular
Class4=CGetControleCapa

[DB:CParametro]
DB=1
DBType=ODBC
ColumnCount=19
Column1=[DataProcessamento], 4, 4
Column2=[CargaSybase], 1, 1
Column3=[AgenciaCentral], 2, 6
Column4=[ValorInferior], 3, 21
Column5=[ControleQualidade], -7, 1
Column6=[ProxImagem], 2, 7
Column7=[ProxCaixa], 5, 2
Column8=[ValorAlcada_Mal], 3, 21
Column9=[ValorAlcada_Env], 3, 21
Column10=[ValorCompensa_Mal], 3, 21
Column11=[ValorCompensa_Env], 3, 21
Column12=[ValorAlcadaDep_Mal], 3, 21
Column13=[ValorAlcadaDep_Env], 3, 21
Column14=[ValorAjusteAuto_Mal], 3, 21
Column15=[ValorAjusteAuto_Env], 3, 21
Column16=[ValorAjusteVincManual_Mal], 3, 21
Column17=[ValorAjusteVincManual_Env], 3, 21
Column18=[PrazoVencimento_Mal], 5, 2
Column19=[PrazoVencimento_Env], 5, 2

[CLS:CGetDataProc]
Type=0
HeaderFile=GetDataProc.h
ImplementationFile=GetDataProc.cpp
BaseClass=CRecordset
Filter=N
VirtualFilter=r
LastObject=CGetDataProc

[DB:CGetDataProc]
DB=1
DBType=ODBC
ColumnCount=19
Column1=[DataProcessamento], 4, 4
Column2=[CargaSybase], 1, 1
Column3=[AgenciaCentral], 2, 6
Column4=[ValorInferior], 3, 21
Column5=[ControleQualidade], -7, 1
Column6=[ProxImagem], 2, 7
Column7=[ProxCaixa], 5, 2
Column8=[ValorAlcada_Mal], 3, 21
Column9=[ValorAlcada_Env], 3, 21
Column10=[ValorCompensa_Mal], 3, 21
Column11=[ValorCompensa_Env], 3, 21
Column12=[ValorAlcadaDep_Mal], 3, 21
Column13=[ValorAlcadaDep_Env], 3, 21
Column14=[ValorAjusteAuto_Mal], 3, 21
Column15=[ValorAjusteAuto_Env], 3, 21
Column16=[ValorAjusteVincManual_Mal], 3, 21
Column17=[ValorAjusteVincManual_Env], 3, 21
Column18=[PrazoVencimento_Mal], 5, 2
Column19=[PrazoVencimento_Env], 5, 2

[CLS:CGetValoresParametro]
Type=0
HeaderFile=GetValoresParametro.h
ImplementationFile=GetValoresParametro.cpp
BaseClass=CRecordset
Filter=N
VirtualFilter=r
LastObject=CGetValoresParametro

[DB:CGetValoresParametro]
DB=1
DBType=ODBC
ColumnCount=19
Column1=[DataProcessamento], 4, 4
Column2=[CargaSybase], 1, 1
Column3=[AgenciaCentral], 2, 6
Column4=[ValorInferior], 3, 21
Column5=[ControleQualidade], -7, 1
Column6=[ProxImagem], 2, 7
Column7=[ProxCaixa], 5, 2
Column8=[ValorAlcada_Mal], 3, 21
Column9=[ValorAlcada_Env], 3, 21
Column10=[ValorCompensa_Mal], 3, 21
Column11=[ValorCompensa_Env], 3, 21
Column12=[ValorAlcadaDep_Mal], 3, 21
Column13=[ValorAlcadaDep_Env], 3, 21
Column14=[ValorAjusteAuto_Mal], 3, 21
Column15=[ValorAjusteAuto_Env], 3, 21
Column16=[ValorAjusteVincManual_Mal], 3, 21
Column17=[ValorAjusteVincManual_Env], 3, 21
Column18=[PrazoVencimento_Mal], 5, 2
Column19=[PrazoVencimento_Env], 5, 2

[CLS:CGetCapaVincular]
Type=0
HeaderFile=GetCapaVincular.h
ImplementationFile=GetCapaVincular.cpp
BaseClass=CRecordset
Filter=N
VirtualFilter=r
LastObject=CGetCapaVincular

[DB:CGetCapaVincular]
DB=1
DBType=ODBC
ColumnCount=23
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
Column16=[QtdDoctos], 4, 4
Column17=[QtdDigitados], 4, 4
Column18=[Conta], 3, 21
Column19=[Dinheiro], 3, 21
Column20=[Diferenca], 3, 21
Column21=[MotivoExclusao], 12, 100
Column22=[Duplicidade], 2, 3
Column23=[Atualizacao], -2, 8

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
ColumnCount=4
Column1=[DataProcessamento], 4, 4
Column2=[IdCapa], 4, 4
Column3=[Comentario], 12, 60
Column4=[IdModulo], 4, 4

