select * from PROCESSADOR
select * from PLACA_MAE
select * from MEMORIA
select * from TIPO_MEMORIA
select * from SOCKET
select * from MARCA_PROCESSADOR
select * from ARMAZENAMENTO
select * from GABINETE
select * from FONTE
select * from PARAMETROS

select	PROCE.DESC_PROCESSADOR , PROCE.FORNECEDOR_NACIONAL , PROCE.VALOR_NACIONAL , PROCE.FORNECEDOR_IMPORTADO , PROCE.VALOR_IMPORTADO
from	PROCESSADOR PROCE, PLACA_MAE PM
where	PROCE.ID_SOCKET = PM.ID_SOCKET

select	PROCE.DESC_PROCESSADOR, PROCE.FORNECEDOR_NACIONAL , PROCE.VALOR_NACIONAL , PROCE.FORNECEDOR_IMPORTADO , PROCE.VALOR_IMPORTADO,
		PM.MODELO_PLACA,PM.FORNECEDOR_NACIONAL , PM.VALOR_NACIONAL , PM.FORNECEDOR_IMPORTADO , PM.VALOR_IMPORTADO
from	PROCESSADOR PROCE, PLACA_MAE PM
where	PROCE.ID_SOCKET = PM.ID_SOCKET
