
#if !defined(STRUCTS_H)
#define STRUCTS_H

typedef struct
{
	long IdCapa;
	CString IdEnv_Mal;
	CString Agencia;
	CString Conta;
	CString Status;
	CString Num_Malote;
	BOOL NovaRegra;
	BOOL GerarAjuste;
	BOOL CSP;
	__int64 Diferenca;

} CAPA;

typedef struct
{
	long IdDocto;
	__int64 Valor;
	int TipoDocto;
	long Vinculo;
	CString TipoGenerico,
		    Leitura,
			Status;
	BOOL Alcada,
		 DesprezarVinculo;
} DOCUMENTO;

#endif