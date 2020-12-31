// VinculoService.cpp: implementation of the CVinculoService class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "VinculoMsg.h"
#include "VinculoService.h"
#include "Vinculador.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CVinculoService::CVinculoService()
:CNTService("MDI_Ubb - Vinculo Automatico")
{
}

CVinculoService::~CVinculoService()
{
}

void CVinculoService::OnContinue()
{
	m_bIsRunning=TRUE;
	LogEvent(EVENTLOG_WARNING_TYPE, EVMSG_CONTINUE);
}

void CVinculoService::OnShutDown()
{
	m_bIsRunning=FALSE;

	//Finaliza o serviço
	TerminateService();
	LogEvent(EVENTLOG_WARNING_TYPE, EVMSG_SHUTDOWN, m_szServiceName);
}

void CVinculoService::OnStop()
{
	m_bIsRunning = FALSE;
}

BOOL CVinculoService::OnInit()
{
	return InitService();
}

void CVinculoService::Run()
{
	int iCodRet;
	DWORD lSleep;

    while (m_bIsRunning) 
	{
		//Processa o vinculo de uma capa
		// retorna: -1 se erro; 0 se nao existe capa; 1 se processou capa
		iCodRet = m_oVinculador.ProcessaVinculo();
		if( iCodRet < 0 ) // Erro
		{
			m_oVinculador.GetLastErrorInfo(m_iCodError, m_MsgError); 
			LogEvent(EVENTLOG_ERROR_TYPE, EVMSG_GENERIC_ERROR, LPCSTR(m_MsgError));
			OnShutDown();
		}
		else if( iCodRet == 0 ) // Nao existe Capa
		{
			lSleep = (DWORD)m_oVinculador.GetSleep();
			if( lSleep > 0 )
				Sleep( lSleep );
			else
				Sleep(POOLING_INTERVAL);
		}
	}
}

void CVinculoService::TerminateService()
{
	//Finaliza o m_oVinclador
	m_oVinculador.Done(); 
}

BOOL CVinculoService::InitService()
{
	/******* Inicializa m_oVinclador *****/
	if (!m_oVinculador.Init())
	{
		m_oVinculador.GetLastErrorInfo(m_iCodError, m_MsgError); 
		LogEvent(EVENTLOG_ERROR_TYPE, EVMSG_GENERIC_ERROR, LPCSTR(m_MsgError));
		return FALSE;
	}

	return TRUE;
}

