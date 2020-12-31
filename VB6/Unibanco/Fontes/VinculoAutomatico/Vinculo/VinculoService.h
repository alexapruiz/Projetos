// VinculoService.h: interface for the CVinculoService class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_VINCULOSERVICE_H__24CBF9EC_4371_11D4_AF4D_000629E201DC__INCLUDED_)
#define AFX_VINCULOSERVICE_H__24CBF9EC_4371_11D4_AF4D_000629E201DC__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "stdafx.h"
#include "VinculoMsg.h"
#include "NTService.h"
#include "Vinculador.h"

// Intervalo para pooling no diretório de solicitação de FAX (em ms)
#define POOLING_INTERVAL  30000

class CVinculoService : public CNTService  
{
public:
	CVinculoService();
	~CVinculoService();
	void  OnContinue();
	void  OnShutDown();
	void  OnStop();
	BOOL  OnInit();
    void  Run();

private:
	int m_iCodError;
	char m_MsgError[256];
	CVinculador m_oVinculador;
	void TerminateService();
	BOOL InitService();
	
};

#endif // !defined(AFX_VINCULOSERVICE_H__24CBF9EC_4371_11D4_AF4D_000629E201DC__INCLUDED_)
