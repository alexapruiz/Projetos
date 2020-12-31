/**********************************************************
      Módulo: CNTService
   
	  Descrição: Classe que encapsula métodos básicos para 
	             codificação de serviços NT.

				Copiado do MSDN Library - Visual Studio 6

*********************************************************/

#if !defined(AFX_NTSERVICE_H__3566D983_8045_11D3_9B4A_204C4F4F5020__INCLUDED_)
#define AFX_NTSERVICE_H__3566D983_8045_11D3_9B4A_204C4F4F5020__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "StdAfx.h"
#include "VinculoMsg.h"
#define SERVICE_CONTROL_USER 128
//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
//Diretiva para gerar log em todos os comandos
//#define APP_LOG 0
//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

class CNTService
{
public:
    CNTService(const char* szServiceName);
    virtual ~CNTService();

	//Fullpath do aplicativo
	char m_szFilePath[_MAX_PATH];    

    BOOL IsInstalled();
    BOOL Install();
    BOOL Uninstall();
    void LogEvent(WORD wType, DWORD dwID,
                  const char* pszS1 = NULL,
                  const char* pszS2 = NULL,
                  const char* pszS3 = NULL);
    BOOL StartService();
    void SetStatus(DWORD dwState);
    BOOL Initialize();
    virtual void Run();
	virtual BOOL OnInit();
    virtual void OnStop();
    virtual void OnInterrogate();
    virtual void OnPause();
    virtual void OnContinue();
    virtual void OnShutdown();
    virtual BOOL OnUserControl(DWORD dwOpcode);
	virtual BOOL ParseStandardArgs(int argc, char* argv[]);
    void DebugMsg(const char* pszFormat, ...);
    
    // static member functions - utilizadas como CallBack
    static void WINAPI ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv);
    static void WINAPI Handler(DWORD dwOpcode);

    // data members
    char m_szServiceName[64];
    int m_iMajorVersion;
    int m_iMinorVersion;
    SERVICE_STATUS_HANDLE m_hServiceStatus;
    SERVICE_STATUS m_Status;
    BOOL m_bIsRunning;

    // static data
    static CNTService* m_pThis; // nasty hack to get object ptr

private:
    HANDLE m_hEventSource;

protected:
   //Variáveis auxiliares para debug
	#ifdef APP_LOG
		CStdioFile m_LogFile;
	#endif

	CString    m_sLogMsg;

};

#endif // !defined(AFX_NTSERVICE_H__3566D983_8045_11D3_9B4A_204C4F4F5020__INCLUDED_)
