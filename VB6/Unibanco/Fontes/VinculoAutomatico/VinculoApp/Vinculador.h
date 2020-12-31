// Vinculador.h: interface for the CVinculador class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_VINCULADOR_H__D5F5FCEC_45CC_11D4_AF4D_000629E201DC__INCLUDED_)
#define AFX_VINCULADOR_H__D5F5FCEC_45CC_11D4_AF4D_000629E201DC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define  INI_FILE		"Vinculo.ini"
#define  MODULO_CSP		31

#include "Structs.h"

class CGetValoresParametro;
class CGetCapaVincular;
class CGetDocumentos;
class CGetAgContaDeposito;
class CGetIdDoctoAjuste;

class CVinculador  
{
public:
	CVinculador();
	virtual ~CVinculador();
	BOOL Init( void );
	void Done( void );
	void GetLastErrorInfo(int& iCodError, char *lpszMsgError);
	// retorna: -1 se erro; 0 se nao existe capa; 1 se existe capa
	int ProcessaVinculo( void );
	long GetSleep( void );

protected:
	CDatabase             m_oDB;
	CGetValoresParametro *m_pParametro;
	CGetCapaVincular     *m_pCapa;
	CGetDocumentos       *m_pDoc;
	CGetAgContaDeposito  *m_pDep;
	CGetIdDoctoAjuste    *m_pAjuste;
	int                   m_iCodError;
	CAPA                  m_Capa;
	char                  m_MsgError[256];
	long                  m_lDataProc;
	int                   m_QtdChequePagto;
	int                   m_QtdContas;
	int                   m_iSleep;
	CArray<DOCUMENTO, DOCUMENTO> m_ArrayDoc;
	MSG                   m_Msg;

	// retorna: -1 se erro; 0 se nao existe capa; 1 se existe capa
	int ObtemCapa( void );
	BOOL LeDataProc( void );
	BOOL LeParametros( void );
	// retorna: -1 se erro; 0 se nao existe documentos; 1 se existe documentos
	int ObtemDocumentos( void );
	BOOL CapaCSP( void );
	BOOL AtualizaDocumentos( void );
	BOOL VerificaSituacao( void );
	BOOL AtualizaStatusCapa( void );
	// retorna: -1 se erro; 1 se Ok
	int VinculaLancamentoInterno( void );
	// retorna: -1 se erro; 1 se Ok
	int VinculaDeposito( void );
	BOOL ObtemAgContaDeposito( long IdDocto, long TipoDocto, int &Ag, CString &Cta );
	BOOL RemoveAjustes( void );
	BOOL InsereAjuste( int Tipo, int Ag, CString Cta, __int64 Valor, long &IdDocto );
    // retorna: -1 se erro; 0 se nao vinculou; 1 se vinculou
	int VinculaDocumentoMalote( void );
	int VinculaDocumentoMaloteRegraAntiga( void );
	int VinculaDocumentoMaloteRegraNova( void );
	void VinculaDocumentoEnvelope( void );
	BOOL DespachaCapa( void );
	// multiplica por 100 e converte
	__int64 Converte( CString Valor );
	BOOL InsereLog( short Acao );
	BOOL PossuiDocumentoTransmitido( void );
	int  NovaRegraLancamentoInterno( void );
	BOOL InsereControleCapa( CString );
	BOOL ContemAjusteVinculo( long );
	void RemoveVinculo( long );
	long VerificaVinculoMalote( void );
};

#endif // !defined(AFX_VINCULADOR_H__D5F5FCEC_45CC_11D4_AF4D_000629E201DC__INCLUDED_)
