// Vinculador.cpp: implementation of the CVinculador class.
//
//////////////////////////////////////////////////////////////////////

/*
  Data      Versao      Descricao
17/11/2000  2.1         Tratamento do Lancamento Interno (tipodocto = 41)
04/01/2001  2.2         Alcada para Cheque Compensado CP (6) e LI (41)
22/01/2001  2.2			Verificar e enviar capas para Conferencia de Ag/Conta
02/03/2001  2.3         Alteracao na conexao sem usar ODBC
12/03/2001  2.4         Nao fazer Ajuste Contabil para LI (tipodocto = 41)
16/04/2001  2.5         Aceitar Cheque 230 (Bandeirante) como Cheque Ubb
06/06/2001  2.6         Nova Regra do Lançamento Interno
20/07/2001  2.7         Considerar Titulos do Bandeirantes que devem ser tratados como Titulos de outros bancos
17/09/2001  2.8         Acerto na rotina DespachaCapa, não estava calculando corretamente a diferenca da Capa
22/02/2002  2.9         Inserção da Rotina InsereControleCapa
22/03/2002  2.10        Contemplação de capas com origem CSP
22/05/2002  2.11        Implementação de títulos de outros bancos com cheque de outros bancos, envio para Vinculo manual
03/10/2002  2.12		Implementação de Chq 3o. com titulos unibanco - não permite chq 3o. com concessionarias (Envelopes)
10/10/2002  2.13		Troca de senha de acesso ao banco de Dados
05/11/2002  2.14		Bloqueio de vinculo para mais de 200 documentos
28/11/2002  2.15		Alterações MDI (FASE 6 - 2002)
11/12/2002  2.16		Acerto de ajuste na Nova Regra LI (FASE 6 - 2002) [Bug]
10/01/2003  2.17		Acerto da Vincula Deposito (FASE 6 - 2002) [Bug]
15/01/2003  2.18		Inclui envio de Capa para Pesquisa de IPVA qdo nao complementado e obtencao da senha do arquivo .ini
31/01/2003  2.19		Acerto do bloqueioo capa com mais de 200 doctos permitir capas apenas com deposito e enviar apenas p/ ilegiveis
*/

#include "stdafx.h"
#include "math.h"
#include "string.h"
#include "Vinculador.h"
#include "GetDataProc.h"
#include "GetValoresParametro.h"
#include "GetCapaVincular.h"
#include "GetDocumentos.h"
#include "GetAgContaDeposito.h"
#include "GetIdDoctoAjuste.h"
#include "GetDocOcorrencia.h"
#include "GetDocumentoTransmitido.h"
#include "GetControleCapa.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CVinculador::CVinculador()
{
}

CVinculador::~CVinculador()
{

}

BOOL CVinculador::Init( void )
{
	char strServer[256];
	char strDatabase[256];
	char strFileName[256];
	char strUsuario[256];
	char strSenha[256];
	CString strConnection;

	if( GetWindowsDirectory( strFileName, 255) <= 0 )
	{
		strcpy(m_MsgError, "Nao foi possivel obter o diretorio do Windows");
		m_iCodError = 501;
		return FALSE;
	}
	strcat(strFileName, "\\");
	strcat(strFileName, INI_FILE);

	if( GetPrivateProfileString( "Vinculo", "Servidor", "", strServer, 255, strFileName) <= 0 )
	{
		strcpy(m_MsgError, "Nao foi possivel obter o nome do Servidor do banco");
		m_iCodError = 502;
		return FALSE;
	}
	
	if( GetPrivateProfileString( "Vinculo", "DataBase", "", strDatabase, 255, INI_FILE ) <= 0 )
	{
		strcpy(m_MsgError, "Nao foi possivel obter o nome do Banco de Dados");
		m_iCodError = 503;
		return FALSE;
	}

	if( GetPrivateProfileString( "Vinculo", "Usuario", "", strUsuario, 255, INI_FILE ) <= 0 )
	{
		strcpy(m_MsgError, "Nao foi possivel obter o Nome do Usuario");
		m_iCodError = 503;
		return FALSE;
	}

	if( GetPrivateProfileString( "Vinculo", "Senha", "", strSenha, 255, INI_FILE ) <= 0 )
	{
		strcpy(m_MsgError, "Nao foi possivel obter a Senha de Acesso ao Banco de Dados");
		m_iCodError = 503;
		return FALSE;
	}
	
	strConnection.Format( "driver={SQL Server};Server=%s;UID=i;PWD=%s;Database=%s;provider=sqloledb",
						  strServer, strSenha, strDatabase );
	try
	{
		m_oDB.OpenEx( LPCTSTR(strConnection), CDatabase::noOdbcDialog );
	}
	catch (CDBException *E)
	{
		//Armazena a informação sobre o erro
		strcpy(m_MsgError, LPCSTR(E->m_strError)); 
		strcat(m_MsgError, " (CVinculador.Init) ");
		m_iCodError= E->m_nRetCode;  

		E->Delete(); 
		return FALSE;		
	}

	m_pParametro = new CGetValoresParametro(&m_oDB);
	m_pCapa = new CGetCapaVincular(&m_oDB);
	m_pDoc = new CGetDocumentos(&m_oDB);
	m_pDep = new CGetAgContaDeposito(&m_oDB);
	m_pAjuste = new CGetIdDoctoAjuste(&m_oDB);

	m_iSleep = 20;
	memset(&m_Msg, '\0', sizeof(MSG));

	return TRUE; 

}

void CVinculador::Done( void )
{

	if( m_pParametro != NULL )
	{
		if ( m_pParametro->IsOpen() )
			m_pParametro->Close();
		delete m_pParametro;
	}
	
	if( m_pCapa != NULL )
	{
		if ( m_pCapa->IsOpen() )
			m_pCapa->Close();
		delete m_pCapa;
	}

	if( m_pDoc != NULL )
	{
		if ( m_pDoc->IsOpen() )
			m_pDoc->Close();
		delete m_pDoc;
	}

	if( m_pDep != NULL )
	{
		if ( m_pDep->IsOpen() )
			m_pDep->Close();
		delete m_pDep;
	}

	if( m_pAjuste != NULL )
	{
		if ( m_pAjuste->IsOpen() )
			m_pAjuste->Close();
		delete m_pAjuste;
	}

	m_oDB.Close(); 
	m_ArrayDoc.RemoveAll();

}

void CVinculador::GetLastErrorInfo(int& iCodError, char *lpszMsgError)
{
	iCodError = m_iCodError;
	strcpy(lpszMsgError, m_MsgError);
}

// retorna: -1 se erro; 0 se nao existe capa; 1 se existe capa
int CVinculador::ObtemCapa( void )
{
	CString strNumMalote;

	while( true )
	{
		try
		{
			if ( m_pCapa->IsOpen() )
				m_pCapa->Close();

			m_pCapa->m_DataProc = m_lDataProc;
			if( !m_pCapa->Open( CRecordset::snapshot, NULL ) )
			{
				//Armazena a informação sobre o erro
				strcpy(m_MsgError, "Erro na obtenção da Capa para Vincular. (CGetCapaVincular.Open)"); 
				m_iCodError= 100;  
				return -1;
			}
			else
			{
				if( m_pCapa->IsEOF() )
				{
					m_pCapa->Close();
					m_pCapa->m_DataProc = m_lDataProc;
					if( !m_pCapa->Open( CRecordset::snapshot, _T("{ Call VA_GetCapaVincular( ? ) }") ) )
					{
						//Armazena a informação sobre o erro
						strcpy(m_MsgError, "Erro na obtenção da Capa para Vincular. (CGetCapaVincular.Open)"); 
						m_iCodError= 100;  
						return -1;
					}
					else
					{
						if( m_pCapa->IsEOF() )
						{
							return 0;
						}
						else
						{
							// sai do while
							break;
						}
					}
				}
				else
				{
					// sai do while
					break;
				}
			}
		}
		catch (CDBException *E)
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, LPCSTR(E->m_strError)); 
			strcat(m_MsgError, " (CGetCapaVincular.Open) ");
			m_iCodError= E->m_nRetCode;  
			E->Delete(); 

			if( memcmp(m_MsgError, "Timeout expired", 15) != 0 )
			{
				return -1;		
			}
		}
	} // while

	m_Capa.IdCapa    = m_pCapa->m_IdCapa;
	m_Capa.IdEnv_Mal = m_pCapa->m_IdEnv_Mal;
	m_Capa.NovaRegra = FALSE;
	m_Capa.Diferenca = 0;

	if( m_Capa.IdEnv_Mal == "M" )
	{
		strNumMalote.Format("%011.11s", LPCTSTR(m_pCapa->m_NumMalote));
		m_Capa.Num_Malote = strNumMalote;
		m_Capa.Agencia = strNumMalote.Left(4);
		m_Capa.Conta   = strNumMalote.Right(7);

		if( strNumMalote.Left(1) == "9" )
			m_Capa.NovaRegra = TRUE;
	}
	else
	{
		m_Capa.Agencia = "";
		m_Capa.Conta   = "";
	}
	m_Capa.Status      = "";
	if( m_pCapa->m_Status == "9" )
	{
		m_Capa.GerarAjuste = TRUE;
	}
	else
	{
		m_Capa.GerarAjuste = FALSE;
	}

	m_pCapa->Close();
	return 1;
}

BOOL CVinculador::LeDataProc( void )
{
	CGetDataProc m_oGetDataProc(&m_oDB);

	try
	{
		if( !m_oGetDataProc.Open( CRecordset::snapshot, NULL ) )
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, "Erro na obtenção da Data de Processamento. (CGetDataProc.Open)"); 
			m_iCodError= 100;  
			return FALSE;
		}
		else
		{
			if( m_oGetDataProc.IsEOF() )
			{
				//Armazena a informação sobre o erro
				strcpy(m_MsgError, "Erro. Não foi possível obter a data do processamento. (CGetDataProc.Eof)"); 
				m_iCodError= 100;  
				return FALSE;
			}
			else
			{
				m_lDataProc = m_oGetDataProc.m_DataProc;
				//m_lDataProc = 20030202;
				m_iSleep    = m_oGetDataProc.m_Sleep;
				m_oGetDataProc.Close();
			}
		}
	}
	catch (CDBException *E)
	{
		//Armazena a informação sobre o erro
		strcpy(m_MsgError, LPCSTR(E->m_strError)); 
		strcat(m_MsgError, " (CGetDataProc.Open) ");
		m_iCodError= E->m_nRetCode;  

		E->Delete(); 
		return FALSE;		
	}
	return TRUE;
}

BOOL CVinculador::LeParametros( void )
{
	try
	{
		if ( m_pParametro->IsOpen() )
			m_pParametro->Close();

		m_pParametro->m_DataProc = m_lDataProc;
		if( !m_pParametro->Open( CRecordset::snapshot, NULL ) )
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, "Erro na obtenção dos Valores de Parametro. (CGetValoresParametro.Open)"); 
			m_iCodError= 100;  
			return FALSE;
		}
		else if( m_pParametro->IsEOF() )
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, "Erro na obtenção dos Valores de Parametro. (CGetValoresParametro.Eof)"); 
			m_iCodError= 100;  
			return FALSE;
		}
	}
	catch (CDBException *E)
	{
		//Armazena a informação sobre o erro
		strcpy(m_MsgError, LPCSTR(E->m_strError)); 
		strcat(m_MsgError, " (CGetValoresParametro.Open) ");
		m_iCodError= E->m_nRetCode;  

		E->Delete(); 
		return FALSE;		
	}
   	
	return TRUE; 
}

// retorna: -1 se erro; 0 se nao existe capa; 1 se existe capa
int CVinculador::ProcessaVinculo( void )
{
	int iCodRet;
	HANDLE hThread;
	BOOL bChangePriority;

	if( !LeDataProc() )
		return -1;

	iCodRet = ObtemCapa();
	if( iCodRet < 1 )
		return iCodRet;

	if( PossuiDocumentoTransmitido() )
	{
		// Gravar Status da Capa 
		if( !AtualizaStatusCapa() )
			return -1;
		else
			return 1;
	}

	/*============================
	  Não pode remover ajustes 
	  de capas que vieram de CSP
	============================*/
	if ( ! CapaCSP() )
		RemoveAjustes();

	iCodRet = ObtemDocumentos();
	if( iCodRet < 0 )
		return iCodRet;
	
	if( iCodRet == 0 ) // Nao existem documentos na Capa ou
					   // Existem documentos para Conf. Ag/Conta
	{
		// Gravar Status da Capa 
		if( !AtualizaStatusCapa() )
			return -1;
		else
			return 1;
	}

	if( !LeParametros() )
		return -1;

	if( !VerificaSituacao() ) // Algum problema na capa
	{
		// Gravar Status da Capa 
		if( !AtualizaStatusCapa() )
			return -1;
		else
			return 1;
	}

	// Se houverem mais de 100 documentos no Envelpe / Malote
	// diminui a prioridade da Thread para o consumo de CPU
	// nao impactar o SQL Server
	if( m_ArrayDoc.GetUpperBound() >= 100 )
	{
		hThread = GetCurrentThread();
    	bChangePriority = SetThreadPriority(hThread, THREAD_PRIORITY_BELOW_NORMAL);
		//bChangePriority = SetThreadPriority(hThread, THREAD_PRIORITY_ABOVE_NORMAL);
	}

	if( VinculaLancamentoInterno() < 0 )
	{
		// Retorna a Thread a prioridade Normal
		if( bChangePriority )
			SetThreadPriority(hThread, THREAD_PRIORITY_NORMAL);
		return -1;
	}

	if( VinculaDeposito() < 0 )
	{
		// Retorna a Thread a prioridade Normal
		if( bChangePriority )
			SetThreadPriority(hThread, THREAD_PRIORITY_NORMAL);
		return -1;
	}

	if( m_Capa.IdEnv_Mal == "M") // Malote
	{
		if( VinculaDocumentoMalote() < 0 )
		{
			// Retorna a Thread a prioridade Normal
			if( bChangePriority )
				SetThreadPriority(hThread, THREAD_PRIORITY_NORMAL);
			return -1;
		}
	}
	else                         // Envelope
	{
		VinculaDocumentoEnvelope();
	}
	
	if( !AtualizaDocumentos() )
	{
		// Retorna a Thread a prioridade Normal
		if( bChangePriority )
			SetThreadPriority(hThread, THREAD_PRIORITY_NORMAL);
		return -1;
	}

	// Retorna a Thread a prioridade Normal
	if( bChangePriority )
		SetThreadPriority(hThread, THREAD_PRIORITY_NORMAL);

	if( !DespachaCapa() )
		return -1;

	return 1;
}

long CVinculador::GetSleep( void )
{
	return m_iSleep * 1000;
}

// retorna: -1 se erro; 0 se nao existe documentos; 1 se existe documentos
int CVinculador::ObtemDocumentos( void )
{
	DOCUMENTO Doc;
	CGetDocOcorrencia m_oDocOcorrencia(&m_oDB);

	m_ArrayDoc.RemoveAll();

	while( true )
	{
		try
		{
			if ( m_pDoc->IsOpen() )
				m_pDoc->Close();

			m_pDoc->m_DataProc = m_lDataProc;
			m_pDoc->m_IdCapa   = m_Capa.IdCapa;
			if( !m_pDoc->Open( CRecordset::snapshot, NULL ) )
			{
				//Armazena a informação sobre o erro
				strcpy(m_MsgError, "Erro na obtenção dos Documentos. (CGetDocumentos.Open)"); 
				m_iCodError= 100;  
				return -1;
			}
			else if( m_pDoc->IsEOF() )
			{
				// Capa sem documentos
				try
				{
					if( m_oDocOcorrencia.IsOpen() )
						m_oDocOcorrencia.Close();

					m_oDocOcorrencia.m_DataProc = m_lDataProc;
					m_oDocOcorrencia.m_IdCapa   = m_Capa.IdCapa;
					if( !m_oDocOcorrencia.Open( CRecordset::snapshot, NULL ) )
					{
						//Armazena a informação sobre o erro
						strcpy(m_MsgError, "Erro na obtenção na Qtde de Documentos com Ocorrencia. (CGetDocOcorrencia.Open)"); 
						m_iCodError= 100;  
						return -1;
					}
					if( m_oDocOcorrencia.m_Qtde > 0)
					{
						// So exitem documentos com ocorrencia na capa
						// Enviar para Transmissao
						m_Capa.Status = "R";
					}
					else
					{
						// Efetivamente nao existem documentos na capa
						// Enviar para Ilegivies
						m_Capa.Status = "5";
					}
					m_pDoc->Close();
					return 0;
				}
				catch(CDBException *E)
				{
					//Armazena a informação sobre o erro
					strcpy(m_MsgError, LPCSTR(E->m_strError)); 
					strcat(m_MsgError, " (CGetDocOcorrencia.Open) ");
					m_iCodError= E->m_nRetCode;  
					E->Delete(); 

					if( memcmp(m_MsgError, "Timeout expired", 15) != 0 )
					{
						return -1;		
					}
					else
						continue;
				}
			}
			break; // sai do while
		}
		catch (CDBException *E)
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, LPCSTR(E->m_strError)); 
			strcat(m_MsgError, " (CGetDocumentos.Open) ");
			m_iCodError= E->m_nRetCode;  
			E->Delete(); 

			if( memcmp(m_MsgError, "Timeout expired", 15) != 0 )
			{
				return -1;		
			}
		}
	} // while
   	
	while( !m_pDoc->IsEOF() )
	{
		Doc.IdDocto          = m_pDoc->m_IdDocto;
		Doc.TipoDocto        = m_pDoc->m_TipoDocto;
		Doc.Valor            = Converte(m_pDoc->m_Valor);
		Doc.Leitura          = m_pDoc->m_Leitura;
		Doc.Status			 = m_pDoc->m_Status;
		/*=================================
		  Se capa veio de CSP
		  Guardar vinculo, pois será usado
		=================================*/
		Doc.Vinculo          = ( m_Capa.CSP == TRUE ? m_pDoc->m_Vinculo : 0 ) ;
		Doc.Alcada           = FALSE;
		Doc.DesprezarVinculo = FALSE;
		
		if( Doc.Status == "L" )
		{
			// Exitem documentos com status = "L" (para Conf. de Ag/Conta)
			// Enviar para Conferencia de Ag/Conta
			m_pDoc->Close();
			m_Capa.Status = "L";
			return 0;
		}

		switch(Doc.TipoDocto)
		{
			case 2:
			case 3:
			case 37:
				// DEPOSITO
				Doc.TipoGenerico = "DE";
				break;
			case 4:
			case 5:
			case 6:
			case 41:
				// ADCC ou CHEQUE DE PAGAMENTO
				Doc.TipoGenerico = "CP";
				break;
			case 7:
				// CHEQUE DE DEPOSITO
				Doc.TipoGenerico = "CD";
				break;
			case 32:
			case 34:
				// ACERTO DE CREDITO
				Doc.TipoGenerico = "AC";
				break;
			case 33:
			case 38:
				// ACERTO DE DEBITO
				Doc.TipoGenerico = "AD";
				break;
			case 39:
				// CAPA DE OCT
				Doc.TipoGenerico = "OC";
				break;
			default:
				// CONTAS
				Doc.TipoGenerico = "CO";
		}

		m_ArrayDoc.Add(Doc);
		m_pDoc->MoveNext();
	}
	
	m_pDoc->Close();
	return 1; 
}

BOOL CVinculador::VerificaSituacao( void )
{
	DOCUMENTO   Doc, DocAux;
	int	        i;
	int			j;
	int			iQtdDepositos    = 0;
	BOOL		Deposito	     = FALSE;
	BOOL		Debito		     = FALSE;
	BOOL		Credito		     = FALSE;
	BOOL		NovaRegraLI      = FALSE;
	BOOL        Ilegiveis        = FALSE;
	BOOL        ContemCPTerceiro = FALSE;
	BOOL        ContemCOTerceiro = FALSE;

	m_QtdChequePagto = 0;
	m_QtdContas      = 0;
	NovaRegraLI      = FALSE;

	// Se capa possuir mais de 200 docto não processar exceto capas apenas com dep/cheque deposito
	if ( m_ArrayDoc.GetUpperBound() > 200 )
	{

		for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
		{
			//Verifica se capa contem apenas depositos permitindo o vinculo
			if ( m_ArrayDoc[i].TipoDocto == 5 || m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
			{
				m_QtdChequePagto ++;

			}
			else if ( m_ArrayDoc[i].TipoDocto == 2 || m_ArrayDoc[i].TipoDocto == 37  )
			{
				iQtdDepositos ++;
			}

			else if	( m_ArrayDoc[i].TipoDocto != 39  )
			{
				m_QtdContas ++;
			}

		}
		if ( ! (m_QtdChequePagto > 0 && iQtdDepositos > 0 && m_QtdContas == 0 ))
		{
			
			m_Capa.Status = "5";
			return FALSE;

		}
	}
	

	//Verifica se capa contem documento IPVA enviando-o para pesquisa de Renavam se nao complementado.
	if ( !m_Capa.GerarAjuste )
	{
		for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
		{
			if ( m_ArrayDoc[i].TipoDocto == 46 && m_ArrayDoc[i].Status == "0" )
			{
				//Seta capa para pesquisa de IPVA
				m_Capa.Status = "*";
				return FALSE;
			}

		}
	}


	if( m_Capa.IdEnv_Mal == "M" && !m_Capa.NovaRegra && m_lDataProc > m_pParametro->m_DataFinalRegraAntiga )
	{
        /**************************************
        ' * Enviando Capa para Ilegiveis      *
        ' *************************************/
		m_Capa.Status = "5";
		return FALSE; // Capa de malote com regra antiga
	}

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];

		if( Doc.TipoDocto == 0 || (Doc.TipoDocto != 39 && Doc.Valor <= 0) )
		{
            /**************************************
            ' * Documento não foi complementado   *
            ' * Enviando Capa para Ilegiveis      *
            ' *************************************/
			m_Capa.Status = "5";
			break;
		}
		else if( Doc.TipoDocto == 41 && m_Capa.IdEnv_Mal == "E" )
		{
            /**************************************
            ' * Nao pode haver LI em Envelope     *
            ' * Enviando Capa para Ilegiveis      *
            ' *************************************/
			m_Capa.Status = "5";
			break;
		}
		else if( Doc.TipoDocto == 39 )
		{
            /**********************
            ' * Capa possui OCT   *
            ' *********************/
			if( m_Capa.IdEnv_Mal != "M" )
			{
				/************************************************
				' * Nao pode haver capa de OCT em envelope      *
				' * Enviando Capa para Ilegiveis                *
				' ***********************************************/
				m_Capa.Status = "5";
				break;
			}
			else if( i == m_ArrayDoc.GetUpperBound() )
			{
				/*******************************************************
				' * OCT não pode ser o último Documento da Capa        *
				' * Enviando Capa para Ilegiveis                       *
				' ******************************************************/
				m_Capa.Status = "5";
				break;
			}
			else if( m_ArrayDoc[i+1].TipoDocto != 37 )
			{
                /********************************************************
                ' * Não pode haver uma capa de OCT sem uma Ficha de OCT *
                ' * Enviando Capa para Ilegiveis                        *
                ' *******************************************************/
				m_Capa.Status = "5";
				break;
			}
		}
		else if( Doc.TipoGenerico == "DE" )
		{
			/***************************
			' * Capa possui Depósito   *
			' **************************/
			Deposito = TRUE;

			if( i == m_ArrayDoc.GetUpperBound() )
			{
				/*******************************************************
				' * Depósito não pode ser o último Documento da Capa   *
				' * Enviando Capa para Ilegiveis                       *
				' ******************************************************/
				m_Capa.Status = "5";
				break;
			}

			/****************************************************
			' * Verificação da nova regra do Lançamento Interno *
			****************************************************/
			iQtdDepositos++;
			NovaRegraLI = TRUE; // Se alguma coisa não bater, setar como FALSE
			j = i + 1;

			while( j <= m_ArrayDoc.GetUpperBound() )
			{
				if( m_ArrayDoc[j].TipoGenerico == "DE" ||
					m_ArrayDoc[j].TipoGenerico == "OC" )
				{
					iQtdDepositos++;
					j++;
					continue;
				}
					/*********************************************
					  O próximo(s) documento(s) precisa ser um LI
					**********************************************/
				else if ( m_ArrayDoc[j].TipoGenerico == "CP" && m_ArrayDoc[j].TipoDocto == 41 )
				{
					/******************************
					  Verifica Qtde de LI's
					******************************/
					
					while(j <= m_ArrayDoc.GetUpperBound() && m_ArrayDoc[j].TipoDocto == 41 )
					{
						j ++ ;
					}
					
					/********************************************
					  Verifica então se são cobranças ou outro LI
					*********************************************/
					
					if (NovaRegraLI == TRUE)
					{
						for (j ; j <= m_ArrayDoc.GetUpperBound(); j++)
						{
							if ( m_ArrayDoc[j].TipoGenerico != "CO" &&
								 m_ArrayDoc[j].TipoGenerico != "AD" &&
								 m_ArrayDoc[j].TipoGenerico != "AC" )

							{
								m_Capa.Status = "5";
								NovaRegraLI   = FALSE;
								break;
							}
						}
					}
					
				}
				else
				{
					NovaRegraLI = FALSE;
					break;
				}
				j++;
			}

			/*****************************************
			  Fim da nova regra do Lançamento Interno
			*****************************************/

			if( m_ArrayDoc[i+1].TipoGenerico == "DE" && NovaRegraLI == FALSE)
			{
                /****************************************
                ' * Não pode haver 2 Depósitos seguidos *
                ' * Enviando Capa para Ilegiveis        *
                ' ***************************************/
				m_Capa.Status = "5";
				break;
			}
			else if( Doc.TipoDocto == 37 )
			{
				/***********************************
				' * Se eh uma OCT                  *
				' **********************************/
				if( m_Capa.IdEnv_Mal != "M" )
				{
					/************************************************
					' * Nao pode haver OCT em envelope              *
					' * Enviando Capa para Ilegiveis                *
					' ***********************************************/
					m_Capa.Status = "5";
					break;
				}
				else if( i == 0 || m_ArrayDoc[i-1].TipoDocto != 39)
				{
                    /********************************************************
                    ' * Não pode haver uma Ficha de OCT sem uma Capa de OCT *
					' * Enviando Capa para Ilegiveis                        *
					' *******************************************************/
					m_Capa.Status = "5";
					break;
				}
				else if( m_ArrayDoc[i+1].TipoDocto == 37 )
				{
					/****************************************
					' * Não pode haver 2 Octs seguidas      *
					' * Enviando Capa para Ilegiveis        *
					' ***************************************/
					m_Capa.Status = "5";
					break;
				}
			}
		}
		else if( Deposito )
		{
            /************************************
            ' * Documento depois de um Depósito *
			'									*
			' * 11-06-2001						*
			' * Não pode estar dentro da nova   *
			' * regra do Lançamento Interno     *
            ' ***********************************/
			if( (Doc.TipoGenerico == "CO" || Doc.TipoDocto == 4) && NovaRegraLI == FALSE)
			{
                /**************************************
                ' * Documento fora de Ordem           *
                ' * Enviando Capa para Ilegiveis      *
                ' *************************************/
				m_Capa.Status = "5";
				break;
			}
			else if( Doc.TipoGenerico != "AD" && Doc.TipoGenerico != "AC" &&
				     Doc.TipoGenerico != "OC" && Doc.TipoDocto != 41 && 
					 NovaRegraLI == FALSE)
			{
                /**************************************************
                ' * Transformando Documento em Cheque de Deposito *
                ' *************************************************/
				m_ArrayDoc[i].TipoDocto = 7;
				m_ArrayDoc[i].Alcada = FALSE;
				m_ArrayDoc[i].TipoGenerico = "CD";
			}
		}
		// Verifica Alcada
		if( m_Capa.IdEnv_Mal == "M" ) // Malote
		{
			// Se CH. UBB > ValorAlcada
			if( (m_ArrayDoc[i].TipoDocto == 5 || m_ArrayDoc[i].TipoDocto == 4) && 
				m_ArrayDoc[i].Valor >= Converte(m_pParametro->m_ValorAlcada_Mal) )
			{
				m_ArrayDoc[i].Alcada = TRUE;
			}
			// Se Dep. CC ou Poup > ValorAlcadaDep
			else if( (m_ArrayDoc[i].TipoDocto == 2 || m_ArrayDoc[i].TipoDocto == 3 ) && 
					  m_ArrayDoc[i].Valor >= Converte(m_pParametro->m_ValorAlcadaDep_Mal) )
			{
				m_ArrayDoc[i].Alcada = TRUE;
			}
			// Se Ch. Compensado (CP) ou LI > ValorAlcadaOutros
			else if( (m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 41 ) && 
					  m_ArrayDoc[i].Valor >= Converte(m_pParametro->m_ValorAlcadaOutros_Mal) )
			{
				m_ArrayDoc[i].Alcada = TRUE;
			}
		}
		else                          // Envelope
		{
			// Se CH. UBB > ValorAlcada
			if( (m_ArrayDoc[i].TipoDocto == 5  || m_ArrayDoc[i].TipoDocto == 4) && 
				m_ArrayDoc[i].Valor >= Converte(m_pParametro->m_ValorAlcada_Env) )
			{
				m_ArrayDoc[i].Alcada = TRUE;
			}
			// Se Dep. CC ou Poup > ValorAlcadaDep
			else if( (m_ArrayDoc[i].TipoDocto == 2 || m_ArrayDoc[i].TipoDocto == 3 ) && 
					  m_ArrayDoc[i].Valor >= Converte(m_pParametro->m_ValorAlcadaDep_Env) )
			{
				m_ArrayDoc[i].Alcada = TRUE;
			}
			// Se Ch. Compensado (CP) ou LI > ValorAlcadaOutros
			else if( (m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 41 ) && 
					  m_ArrayDoc[i].Valor >= Converte(m_pParametro->m_ValorAlcadaOutros_Env) )
			{
				m_ArrayDoc[i].Alcada = TRUE;
			}

			j = 0;
			while( j <= m_ArrayDoc.GetUpperBound() )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.TipoGenerico == "CP" && DocAux.TipoDocto == 6 )
				{
					Ilegiveis = TRUE;
				}
				else if( DocAux.TipoGenerico == "CO" && (DocAux.TipoDocto == 20 || DocAux.TipoDocto == 21 || DocAux.TipoDocto == 22 || DocAux.TipoDocto == 23 ))
				{
					Ilegiveis = TRUE;
				}
				else
				{
					Ilegiveis = FALSE;
					break;
				}
				j++;
			}
			if( Ilegiveis == TRUE )
			{
				m_Capa.Status = "5";
				break;
			}
		}

		/*==================================================
		  Se conter Cheque de terceiro e Conta de Terceiro,
		  em um envelope ou Malote Regra Nova, enviar para
		  Vinculo Manual
		==================================================*/
		if( !Deposito && m_Capa.CSP == FALSE && m_Capa.GerarAjuste == TRUE &&
			((m_Capa.IdEnv_Mal == "M" && m_Capa.NovaRegra) || (m_Capa.IdEnv_Mal == "E")) )
		{
			j = 0;
			while( j <= m_ArrayDoc.GetUpperBound() )
			{
				DocAux = m_ArrayDoc[j];

				//junior verificar aqui vinculo cheque 3o.
				if( DocAux.TipoGenerico == "CP" && DocAux.TipoDocto == 6 )
				{
					ContemCPTerceiro = TRUE;
				}
				if( DocAux.TipoGenerico == "CO" && (DocAux.TipoDocto != 28 &&
													DocAux.TipoDocto != 29 &&
													DocAux.TipoDocto != 30))
				{
					ContemCOTerceiro = TRUE;
				}
				j++;
			}

			if( ContemCPTerceiro == TRUE && ContemCOTerceiro == TRUE )
			{
				m_Capa.Status = "7";
				break;
			}
		}

		if( (Doc.TipoDocto >= 4 && Doc.TipoDocto <= 7) || Doc.TipoDocto == 41 )
		{
			Debito = TRUE;
		}
		else if( Doc.TipoDocto != 39 )
		{
			Credito = TRUE;
		}

		if( Doc.TipoGenerico == "CP" )
		{
			m_QtdChequePagto++;
		}
		else if( Doc.TipoGenerico == "CO" )
		{
			m_QtdContas++;
		}

	}
	if( m_Capa.Status == "" && (!Debito || !Credito) )
	{
		/********************************************************
		' * Deve haver pelo menos um Debito e pelo menos um     *
		' * Credito                                             *
		' * Enviando Capa para Ilegiveis                        *
		' *******************************************************/
		m_Capa.Status = "5";
	}

	if( !m_Capa.Status.IsEmpty() )
		return FALSE; // Capa com problemas
	else
		return TRUE; // Capa Ok
}

BOOL CVinculador::AtualizaStatusCapa( void )
{
	CString strSql;

	strSql.Format("Execute VA_AtualizaStatusCapa %ld, %ld, '%s', %.2f", 
				   m_lDataProc, m_Capa.IdCapa, LPCTSTR(m_Capa.Status), ((double)abs(m_Capa.Diferenca) / 100));

	try
	{
		m_oDB.ExecuteSQL(LPCTSTR(strSql));
	}
	catch (CDBException *E)
	{
		//Armazena a informação sobre o erro
		strcpy(m_MsgError, LPCSTR(E->m_strError)); 
		strcat(m_MsgError, " (AtualizarStatusCapa) ");
		m_iCodError= E->m_nRetCode;  

		E->Delete(); 
		return FALSE;		
	}

	if( m_Capa.Status == "4" )
	{
		InsereLog( 110 );
	}
	else if( m_Capa.Status == "5" )
	{
		/*--------------------------------------------------
		  Gravar ControleCapa quando envio é para Ilegiveis
		--------------------------------------------------*/
		InsereControleCapa("");

		InsereLog( 111 );
	}
	else if( m_Capa.Status == "6" )
	{
		InsereLog( 112 );
	}
	else if( m_Capa.Status == "7" )
	{
		InsereLog( 113 );
	}
	else if( m_Capa.Status == "R" )
	{
		InsereLog( 114 );
	}
	else if( m_Capa.Status == "V" )
	{
		InsereLog( 115 );
	}

	return TRUE;
}

// retorna: -1 se erro; 1 se Ok
int CVinculador::VinculaLancamentoInterno( void )
{

	DOCUMENTO	Doc, DocAux, Ajuste;
	int			i, j, IndiceDeposito, TipoAjusteCredito, TipoAjusteDebito, ContaDepositos, ContaContas, QtdeLI;
	long		lVinculo, IdDocto;
	__int64		ValorAjusteContabil, ValorLI, SomaConta, SomaDeposito, Diferenca;
	int			Agencia;
	CString		Conta;
	BOOL		bContemAjuste;

	i				  = 0;
	IndiceDeposito	  = 0;
	Agencia			  = 0;
	TipoAjusteCredito = 0;
	TipoAjusteDebito  = 0;
	ContaDepositos    = 0; // Contagem dos depositos na capa
	ContaContas		  = 0; // Contagem das contas na capa
	QtdeLI			  = 0; // Conta Qtde de LI's


	while( i <= m_ArrayDoc.GetUpperBound() )
	{
		Doc = m_ArrayDoc[i];
		if (Doc.TipoGenerico == "CP" && Doc.TipoDocto == 41 && i > 0)
		{
			lVinculo = Doc.IdDocto;
	        ValorLI = 0;
			SomaConta = 0;
			SomaDeposito = 0;

			bContemAjuste = ContemAjusteVinculo( lVinculo );

			/****************************
			  Soma valores dos Depositos
			****************************/
			j = i - 1;
			if( m_ArrayDoc[j].TipoGenerico == "DE" )
			{
				while( 
						(j >= 0) && 
						(m_ArrayDoc[j].TipoGenerico == "DE" ||
					     m_ArrayDoc[j].TipoGenerico == "OC")
					 )
				{
					SomaDeposito += m_ArrayDoc[j].Valor;
					/**********************************
					  Incrementa contador de depositos
					**********************************/
					ContaDepositos++;

					if( j % 100 == 0 )
					{
						PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
					}
					IndiceDeposito = j;
					j--;
				}
			}
			/**********************************************
			  Se não for deposito, não está dentro da regra
			**********************************************/
			else
				return 1;

			/*****************************
			 Conta e soma os Li's na Capa
			******************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++)
			{
				if( m_ArrayDoc[j].TipoDocto == 41 )
				{	
					ValorLI += m_ArrayDoc[j].Valor;
					QtdeLI ++;
				}
			}

			/*************************
			  Soma valores das contas
			*************************/
			j = i + QtdeLI;
			if( j <= m_ArrayDoc.GetUpperBound() && m_ArrayDoc[j].TipoGenerico == "CO" )
			{
				while( j <= m_ArrayDoc.GetUpperBound() )
				{
					if( m_ArrayDoc[j].TipoGenerico == "CO" )
					{
						SomaConta += m_ArrayDoc[j].Valor;
						ContaContas++;
					}
					else if( m_ArrayDoc[j].TipoGenerico == "AC" || m_ArrayDoc[j].TipoGenerico == "AD" ) // para tratamento de capas vindas de CSP
					{
						j++;
						continue;
					}
					else
						return -1;

					if( j % 100 == 0 )
					{
						PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
					}
					j++;
				}
			}


			if (lVinculo == 0)
				return -1;

			Diferenca = ValorLI - (SomaConta + SomaDeposito);
			i = j;

			/**********************************************
			  Se diferenca das contas e Lançamento Interno
			  forem menor do que está na tabela parametro
			  ValorAjuste, gerar ajuste
			**********************************************/
			if( ( abs(Diferenca) > 0 && m_Capa.GerarAjuste == TRUE ) ||
				( abs(Diferenca) > 0 && m_Capa.CSP == TRUE ) )
			{

				/*********************************
				  Pega o valor do Ajuste Contabil
				*********************************/
				ValorAjusteContabil = Converte(m_pParametro->m_ValorAjusteContabil);

				if( (abs(Diferenca) <= ValorAjusteContabil) ||
					(abs(Diferenca) >  ValorAjusteContabil && m_Capa.NovaRegra == FALSE && ContaDepositos >= 1 && ContaContas == 0) ||
					(abs(Diferenca) >  ValorAjusteContabil && m_Capa.NovaRegra == TRUE && ContaDepositos == 1 && ContaContas == 0))
				{

					/********************************************
					  De acordo com a Selma (26-07-2001), quando
					  diferenca for maior que ValorAjusteContabil
					  e for malote antigo. Gerar ajuste na conta do
					  malote
					********************************************/
					TipoAjusteCredito = 42;
					TipoAjusteDebito  = 43;
					Agencia = 0;
					Conta = "0";

					/*****************************
					  Para Malote Novo, com 
					  1 DEPTO
					  1 LANCTO
					  Ajuste vai para o deposito
					*****************************/
					if( m_Capa.NovaRegra == TRUE )
					{
						DocAux = m_ArrayDoc[IndiceDeposito];

						if( ContaDepositos == 1 && ContaContas == 0 && abs(Diferenca) > ValorAjusteContabil)
						{
							TipoAjusteCredito = 32;
							TipoAjusteDebito  = 33;

							/*******************************************
							  Obtem-se a Agencia e Conta do Deposito
							  para que seja efetuado o ajuste na conta
							  do Deposito
							*******************************************/
							if( !ObtemAgContaDeposito( DocAux.IdDocto, DocAux.TipoDocto, Agencia, Conta ) )
								return -1;
						}
					}
					/***************
					  Malote Velho
					***************/
					else
					{
						/******************************************************
						  Definir para onde vai o ajuste
						  Se 1 Depto, Ajuste vai para a conta do Deposito
						  Se 2 ou mais Deptos, Ajuste vai para conta do Malote
						******************************************************/
						if( abs(Diferenca) > ValorAjusteContabil )
						{
							TipoAjusteCredito = 32;
							TipoAjusteDebito  = 33;
						}

						if( ContaDepositos == 1 && ContaContas == 0 )
						{
							DocAux = m_ArrayDoc[IndiceDeposito];

							if( !ObtemAgContaDeposito( DocAux.IdDocto, DocAux.TipoDocto, Agencia, Conta ) )
								return -1;
						}
						else if ( ContaDepositos >= 2 && ContaContas == 0 )
						{
							Agencia = atoi(m_Capa.Agencia);
							Conta = m_Capa.Conta;
						}
					}

					/*****************
					  Gerar o ajuste
					******************/
					if( !(m_Capa.CSP == TRUE && bContemAjuste == TRUE ) )
					{
						if( !InsereAjuste( ( Diferenca < 0 ? TipoAjusteDebito : TipoAjusteCredito),
											Agencia, Conta, abs(Diferenca), IdDocto ))
						{
							return -1;
						}

						Ajuste.Alcada = FALSE;
						Ajuste.DesprezarVinculo = FALSE;
						Ajuste.IdDocto = IdDocto;
						Ajuste.TipoDocto = (Diferenca < 0 ? TipoAjusteDebito : TipoAjusteCredito);
						Ajuste.TipoGenerico = (Diferenca < 0 ? "AD" : "AC");
						Ajuste.Valor = abs(Diferenca);
						Ajuste.Vinculo = 0;

						//m_ArrayDoc.InsertAt(i+1, Ajuste, 1);
						m_ArrayDoc.InsertAt(i, Ajuste, 1);
					}
				}
				else
					break;


			}

			/*******************************
			  Pega apartir do deposito que 
			  fora somado anteriormente
			*******************************/
			
			if( (abs(Diferenca) == 0 ) ||
				(abs(Diferenca) != 0 && m_Capa.GerarAjuste == TRUE && (abs(Diferenca) <= ValorAjusteContabil)) ||
				(abs(Diferenca) > ValorAjusteContabil && m_Capa.NovaRegra == FALSE && m_Capa.GerarAjuste == TRUE ) ||
				(abs(Diferenca) > ValorAjusteContabil && m_Capa.NovaRegra == TRUE  && m_Capa.GerarAjuste == TRUE && ContaDepositos == 1 && ContaContas == 0) ||
				(m_Capa.CSP == TRUE )
			  )
			{
				for ( j = IndiceDeposito; j <= m_ArrayDoc.GetUpperBound(); j++)
				{

					m_ArrayDoc[j].Vinculo = lVinculo;

					if( j % 100 == 0 )
					{
						PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
					}
				}
			}
		}
		i++;
	}
	return 1;
}

// retorna: -1 se erro; 1 se Ok; -2 CSP inválido
int CVinculador::VinculaDeposito( void )
{
	DOCUMENTO Doc       , Ajuste;
	int       i         , j, IndiceLI, v;
	long      lVinculo  , IdDocto;
	__int64   SomaCheque, SomaDeposito, Diferenca;
	int       Agencia;
	BOOL      bContemAjuste, bContemVinculo;
	CString   Conta;


	/*****************************
	  Não faz a vincula deposito 
	  se capa esta na Nova Regra
	  do Lançamento Interno
	*****************************/
	IndiceLI = 0;
	bContemVinculo = FALSE;

	IndiceLI = NovaRegraLancamentoInterno();

	if( IndiceLI == -1 )
		return 1;

	if( IndiceLI == FALSE )
		IndiceLI = m_ArrayDoc.GetUpperBound() + 1;

	// só procura por vinculo
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++)
	{
		if( m_ArrayDoc[i].Vinculo > 0 )
		{
			bContemVinculo = TRUE;
			break;
		}
	}

	i = 0;
	v = 0;
	lVinculo = 0;
	bContemAjuste = FALSE;
	SomaCheque = 0;
	SomaDeposito = 0;

	/************************************************
	  Regra da CSPz

	  Se Capa veio de CSP com vínculo,
	  o ajuste (se for necessário) deverá existir.

	  Se Capa veio de CSP sem vínculo,
	  o Vinculo Automático tem permissão de gerar
	  o ajuste (se for necessário) qual quer que
	  seja o valor. (Somente Depositos)
	************************************************/

	/*==============================
	  Para capa que não veio de CSP
	  tratamento normal, ou veio de
	  CSP, mas não tem vinculo
	==============================*/

	if( (m_Capa.CSP == FALSE) || ( m_Capa.CSP == TRUE && bContemVinculo == FALSE ) )
	{
		while( (i <= m_ArrayDoc.GetUpperBound()) && (i < IndiceLI) )
		{
			Doc = m_ArrayDoc[i];

			lVinculo = 0;

			if( Doc.TipoGenerico == "DE" && 
				Doc.Vinculo == 0 && 
				i < IndiceLI )
			{

				SomaDeposito = Doc.Valor;
				SomaCheque = 0;

				lVinculo = Doc.IdDocto;

				for( j = i + 1; j<= m_ArrayDoc.GetUpperBound() && j < IndiceLI ; j++ )
				{
					if( m_ArrayDoc[j].TipoGenerico == "CD" ||
						m_ArrayDoc[j].TipoDocto == 41 ) // LI
					{
						SomaCheque += m_ArrayDoc[j].Valor;
					}
					else
					{
						break;
					}
					if( j % 100 == 0 )
					{
						PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
					}
				}
				Diferenca = SomaCheque - SomaDeposito;
				if( abs(Diferenca) > 0 )
				{
					if( !m_Capa.GerarAjuste && m_Capa.CSP == FALSE )
					{
						i++;
						continue;
					}

					// Gerar Ajuste Debito / Credito
					if( !ObtemAgContaDeposito( Doc.IdDocto, Doc.TipoDocto, Agencia, Conta ) )
					{
						i++;
						continue;
					}
					if( !InsereAjuste( (SomaCheque > SomaDeposito ? 32 : 33), Agencia, Conta, abs(Diferenca), IdDocto) )
					{
						i++;
						continue;
					}
					Ajuste.Alcada = FALSE;
					Ajuste.DesprezarVinculo = FALSE;
					Ajuste.IdDocto = IdDocto;
					Ajuste.TipoDocto = (SomaCheque > SomaDeposito ? 32 : 33);
					Ajuste.TipoGenerico = (SomaCheque > SomaDeposito ? "AC" : "AD");
					Ajuste.Valor = abs(Diferenca);
					Ajuste.Vinculo = 0;
					m_ArrayDoc.InsertAt(i+1, Ajuste, 1);
					IndiceLI++;
				}
				// Se Oct entao vincular tambem a Capa
				if( Doc.TipoDocto == 37 )
				{
					m_ArrayDoc[i-1].Vinculo   = lVinculo;
				}

				m_ArrayDoc[i].Vinculo = lVinculo;

				// Vincula todos os cheques e ajustes ao deposito/oct
				for( j = i + 1; j <= m_ArrayDoc.GetUpperBound() && j < IndiceLI ; j++ )
				{
					if( m_ArrayDoc[j].TipoGenerico == "CD" ||
						m_ArrayDoc[j].TipoDocto == 32 ||
						m_ArrayDoc[j].TipoDocto == 33 ||
						m_ArrayDoc[j].TipoDocto == 41 )// LI 
					{
						m_ArrayDoc[j].Vinculo   = lVinculo;
					}
					else
					{
						break;
					}
					if( j % 100 == 0 )
					{
						PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
					}
				}
			}
			i++;
		}
	}
	/*=======================================
	  ou se não Tratamento de vinculo de CSP
	=======================================*/
	else
	{
		/*=====================================
		  Tratamento pelo vínculo (se houver)
		=====================================*/
		while( (v <= m_ArrayDoc.GetUpperBound()) && (v < IndiceLI) )
		{
			Doc = m_ArrayDoc[v];

			if( (Doc.Vinculo > 0) && (Doc.Vinculo != lVinculo) && (Doc.TipoGenerico == "DE") )
			{
				lVinculo = Doc.Vinculo;
				i = 0;
				SomaDeposito = 0;
				SomaCheque = 0;
				//Percorre a capa atraz dos documentos com o mesmo vinculo

				while( (i <= m_ArrayDoc.GetUpperBound()) && (i < IndiceLI) )
				{
					Doc = m_ArrayDoc[i];

					if( Doc.Vinculo == lVinculo )
					{
						/*==========================
						  Documento é um Deposito?
						==========================*/
						if( Doc.TipoGenerico == "DE" && Doc.Vinculo != 0 && i < IndiceLI )
						{
							SomaDeposito += Doc.Valor;
						}
						/*================================
						  Documento é um Cheque Deposito
						================================*/
						if( Doc.TipoGenerico == "CD" || Doc.TipoDocto == 41 )
						{
							SomaCheque += Doc.Valor;
						}
						/*=================================
						  Documento é um Acerto de Crédito
						=================================*/
						if( Doc.TipoGenerico == "AC" )
						{
							SomaDeposito += Doc.Valor;
							bContemAjuste = TRUE;
						}
						/*=================================
						  Documento é um Acerto de Crédito
						=================================*/
						if( Doc.TipoGenerico == "AD" )
						{
							SomaCheque += Doc.Valor;
							bContemAjuste = TRUE;
						}
					}
					i++;
					if( i % 100 == 0 )
					{
						PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
					}
				}

				Diferenca = SomaCheque - SomaDeposito;

				if( abs(Diferenca) > 0 )
				{
					/*=================================
					  Capa que contem ajustes e ainda 
					  tem diferenca, voltar para CSP
					=================================*/
					if( bContemAjuste == TRUE )
					{
						//m_Capa.Status = "N";
						//return -2;
						return 1;
					}
					// Gerar Ajuste Debito / Credito
					if( !ObtemAgContaDeposito( Doc.IdDocto, Doc.TipoDocto, Agencia, Conta ) )
					{
						//return -2;
						return 1;
					}
					if( !InsereAjuste( (SomaCheque > SomaDeposito ? 32 : 33), Agencia, Conta, abs(Diferenca), IdDocto) )
					{
						//return -2;
						return 1;
					}
					Ajuste.Alcada = FALSE;
					Ajuste.DesprezarVinculo = FALSE;
					Ajuste.IdDocto = IdDocto;
					Ajuste.TipoDocto = (SomaCheque > SomaDeposito ? 32 : 33);
					Ajuste.TipoGenerico = (SomaCheque > SomaDeposito ? "AC" : "AD");
					Ajuste.Valor = abs(Diferenca);
					Ajuste.Vinculo = 0;
					m_ArrayDoc.InsertAt(i+1, Ajuste, 1);
					IndiceLI++;

					// Se Oct entao vincular tambem a Capa
					if( Doc.TipoDocto == 37 )
					{
						m_ArrayDoc[v-1].Vinculo = lVinculo;
					}

					m_ArrayDoc[v].Vinculo = lVinculo;

					// Vincula todos os cheques e ajustes ao deposito/oct
					for( j = v + 1; j <= m_ArrayDoc.GetUpperBound() && j < IndiceLI ; j++ )
					{
						if( m_ArrayDoc[j].TipoGenerico == "CD" ||
							m_ArrayDoc[j].TipoDocto == 32 ||
							m_ArrayDoc[j].TipoDocto == 33 ||
							m_ArrayDoc[j].TipoDocto == 41 )// LI 
						{
							m_ArrayDoc[j].Vinculo   = lVinculo;
						}
						else
						{
							break;
						}
						if( j % 100 == 0 )
						{
							PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
						}
					}
				}
			}
			v++;
		}
	}
	return 1;
}


BOOL CVinculador::ObtemAgContaDeposito( long IdDocto, long TipoDocto, int &Ag, CString &Cta )
{
	if ( m_pDep->IsOpen() )
		m_pDep->Close();

	m_pDep->m_DataProc  = m_lDataProc;
	m_pDep->m_IdDocto   = IdDocto;
	m_pDep->m_TipoDocto = TipoDocto;
	try
	{
		if( !m_pDep->Open( CRecordset::snapshot, NULL ) )
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, "Erro na obtenção da Ag. e Conta do Deposito. (CGetAgContaDeposito.Open)"); 
			m_iCodError= 100;  
			return FALSE;
		}
		else if( m_pDep->IsEOF() )
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, "Erro na obtenção da Ag. e Conta do Deposito. (CGetAgContaDeposito.Eof)"); 
			m_iCodError= 100;  
			return FALSE;
		}
	}
	catch (CDBException *E)
	{
		//Armazena a informação sobre o erro
		strcpy(m_MsgError, LPCSTR(E->m_strError)); 
		strcat(m_MsgError, " (AtualizarStatusCapa) ");
		m_iCodError= E->m_nRetCode;  

		E->Delete(); 
		return FALSE;		
	}

	Ag  = m_pDep->m_Agencia;
	Cta = m_pDep->m_Conta;

	m_pDep->Close();
	return TRUE;

}

BOOL CVinculador::InsereControleCapa(CString pComentario)
{

	CString sSql;

	if( pComentario == "" )
		pComentario = "Null";

	sSql.Format("Execute InsereControleCapa %ld, %ld, %s, %d",
		         m_lDataProc, m_Capa.IdCapa, LPCTSTR(pComentario), 32);

	try
	{
		m_oDB.ExecuteSQL(LPCTSTR(sSql));
	}
	catch (CDBException *E)
	{
		//Armazena a informação sobre o erro
		strcpy(m_MsgError, LPCSTR(E->m_strError)); 
		strcat(m_MsgError, " (InsereControleCapa) ");
		m_iCodError= E->m_nRetCode;  

		E->Delete(); 
		return FALSE;		
	}

	return TRUE;
}

BOOL CVinculador::InsereAjuste( int Tipo, int Ag, CString Cta, __int64 Valor, long &IdDocto )
{
	CString strSql;

	strSql.Format("Execute VA_InsereAjuste %ld, %ld, %d, %d, %s, %.2f",
				  m_lDataProc, m_Capa.IdCapa, Tipo, Ag, LPCTSTR(Cta), ((double)Valor / 100));

	try
	{
		m_oDB.ExecuteSQL(LPCTSTR(strSql));
	}
	catch (CDBException *E)
	{
		//Armazena a informação sobre o erro
		strcpy(m_MsgError, LPCSTR(E->m_strError)); 
		strcat(m_MsgError, " (InsereAjuste) ");
		m_iCodError= E->m_nRetCode;  

		E->Delete(); 
		return FALSE;		
	}

	try
	{
		m_pAjuste->m_DataProc  = m_lDataProc;
		m_pAjuste->m_IdCapa    = m_Capa.IdCapa;
		m_pAjuste->m_TipoDocto = Tipo;
		m_pAjuste->m_Valor     = ((double)Valor / 100);
		if( !m_pAjuste->Open( CRecordset::snapshot, NULL ) )
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, "Erro na obtenção do IdDocto do Ajuste. (CGetIdDoctoAjuste.Open)"); 
			m_iCodError= 100;  
			return FALSE;
		}
		else
		{
			if( m_pAjuste->IsEOF() )
			{
				//Armazena a informação sobre o erro
				strcpy(m_MsgError, "Erro na obtenção do IdDocto do Ajuste. (CGetIdDoctoAjuste.Eof)"); 
				m_iCodError= 100;  
				return FALSE;
			}
		}
	}
	catch (CDBException *E)
	{
		//Armazena a informação sobre o erro
		strcpy(m_MsgError, LPCSTR(E->m_strError)); 
		strcat(m_MsgError, " (InsereAjuste) ");
		m_iCodError= E->m_nRetCode;  

		E->Delete(); 
		return FALSE;		
	}

	IdDocto = m_pAjuste->m_IdDocto;
	m_pAjuste->Close();

	return TRUE;
}

// retorna: -1 se erro; 0 se nao vinculou; 1 se vinculou
int CVinculador::VinculaDocumentoMalote( void )
{
	if( m_Capa.NovaRegra )
		return VinculaDocumentoMaloteRegraNova();
	else
		return VinculaDocumentoMaloteRegraAntiga();
}

// retorna: -1 se erro; 0 se nao vinculou; 1 se vinculou
int CVinculador::VinculaDocumentoMaloteRegraNova( void )
{
	DOCUMENTO Doc, DocAux, Ajuste;
	int i, j, iQtdCheques, iQtdContas;
	int iQtdSemVinculo, iQtdADCC, iQtdLI, iQtdChPagto, iQtdChTerc;
	long iVinculo, IdDocto, lVinculo;
	__int64 ValorVinculo, Diferenca, ValorCheques, ValorContas;
	int iConta, iCheque;
	int iInicio, iDesprezar;
	BOOL ContemCPTerceiro, ContemCOTerceiro;
	CArray<int, int> aIndConta;
	CArray<int, int> aIndCheque;

	/* Comentario do Vagner:
	   Como a nova regra do Malote e quase tao restritiva quanto a
	   do Envelope, este algoritmo foi uma adaptação do vinculo do
	   Envelope. Porem foi adicionado a ele a possibilidade de vincular
	   N cheques a 1 conta e N cheques a N contas, desde que os cheques
	   sejam todos do unibanco. Alem disso nao e permitido fazer ajustes
	   caso o vinculo envolva mais de um cheque.
	*/

	/*****************************************
    ' * Marcar para Desprezar para o Vínculo *
    ' * Os Cheques com o mesmo Valor         *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		iQtdCheques = 0;
		iQtdContas  = 0;

		Doc = m_ArrayDoc[i];
		
		if( Doc.DesprezarVinculo == FALSE && Doc.Vinculo == 0 && 
			Doc.TipoDocto > 3 && Doc.TipoDocto <= 6 )
		{
            /****************************
            ' * Cheque/ADCC sem Vínculo *
            ' ***************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.Vinculo == 0 && DocAux.Valor == Doc.Valor )
				{
					if( DocAux.TipoDocto > 3 && DocAux.TipoDocto <= 6 )
					{
                        /*******************************
                        ' * Cheque/ADCC no Mesmo Valor *
                        ' ******************************/
						iQtdCheques++;
					}
					else if( DocAux.TipoGenerico == "CO" )
					{
                        /*************************
                        ' * Conta no mesmo Valor *
                        ' ************************/
						iQtdContas++;
					}
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}

		if( iQtdContas > 1 && m_QtdChequePagto > 1 )
		{
            /*************************************************
            ' * Desprezar para o Vínculo Automático          *
            ' ************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdContas > 1 && iQtdCheques > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdCheques > 1 && m_QtdContas > 1 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdCheques > 1 && iQtdContas > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}

		if( m_ArrayDoc[i].DesprezarVinculo )
		{
            /*****************************************
            ' * Marcar para Desprezar para o Vínculo *
            ' * Todos que tenham o mesmo valor       *
            ' ****************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				if( Doc.Valor == m_ArrayDoc[j].Valor )
					m_ArrayDoc[j].DesprezarVinculo = TRUE;
			}
		}
	}

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		iQtdContas     = 0;
		iQtdCheques    = 0;

		Doc = m_ArrayDoc[i];
		if( Doc.DesprezarVinculo == FALSE && Doc.Vinculo == 0 && 
			Doc.TipoGenerico == "CO" )
		{
            /***********************
            ' * Cheque sem Vínculo *
            ' **********************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.Vinculo == 0 && DocAux.Valor == Doc.Valor )
				{
					if( DocAux.TipoDocto > 3 && DocAux.TipoDocto <= 6 )
						iQtdCheques++;
					else if( DocAux.TipoGenerico == "CO" )
						iQtdContas++;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}


		if( iQtdCheques > 1 && m_QtdContas > 1 )
		{
            /*************************************************
            ' * Desprezar para o Vínculo Automático          *
            ' ************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdCheques > 1 && iQtdContas > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdContas > 1 && m_QtdChequePagto > 1 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdContas > 1 && iQtdCheques > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}

		if( m_ArrayDoc[i].DesprezarVinculo )
		{
            /*****************************************
            ' * Marcar para Desprezar para o Vínculo *
            ' * Todos que tenham o mesmo valor       *
            ' ****************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				if( Doc.Valor == m_ArrayDoc[j].Valor )
					m_ArrayDoc[j].DesprezarVinculo = TRUE;
			}
		}
		
	}

    /****************************************
    ' * Primeira Fase do Vinculo            *
    ' * Vinculando Um Cheque para Uma Conta *
    ' ***************************************/

	/* Comentario do Vagner:
	   Nesta fase, se o Cheque de Pagamento for do Unibanco
	   pode-se vincula-lo a qualquer Cobranca/Arrecadacao.
	   Se o Cheque de Pagamento nao for do Unibanco
	   somente vincula-o se a Cobranca for diferente
	   de "Titulos Terceiros Sem CB" (12) e diferente de
	   "Cobranca Terceiros" (31)
	*/

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CP" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /**********************************************
            ' * Cheque, LI ou  ADCC a Verificar o Vinculo *
            ' *********************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.TipoGenerico == "CO" && 
					DocAux.Vinculo == 0 && 
					!DocAux.DesprezarVinculo &&
					Doc.Valor == DocAux.Valor )
				{
					/************************************
					  Versao 2.7
					  Considerar Titulos do Bandeirantes 
					  olhando para o codigo de barras
					************************************/
					if( (Doc.TipoDocto == 6 && DocAux.TipoDocto != 12 && DocAux.TipoDocto != 31) ||
						 Doc.TipoDocto == 5 ||
						 Doc.TipoDocto == 4 ||
						 Doc.TipoDocto == 41 ||
						(DocAux.TipoDocto == 31 && DocAux.Leitura.Left(3) == _T("230")))
					{
						/**********************************
						' * Vinculando Conta com o Cheque *
						' *********************************/
						m_ArrayDoc[i].Vinculo  = Doc.IdDocto;
						m_ArrayDoc[j].Vinculo  = Doc.IdDocto;
						/**************************************
						' * Corrige o tipodocumento do cheque *
						' *************************************/
						if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
						{
							if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230") )
							{
								m_ArrayDoc[i].TipoDocto = 5;
								if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
								{
									m_ArrayDoc[i].Alcada = TRUE;
								}
								else
								{
									m_ArrayDoc[i].Alcada = FALSE;
								}
							}
							else
							{
								m_ArrayDoc[i].TipoDocto = 6;
								if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
								{
									m_ArrayDoc[i].Alcada = TRUE;
								}
								else
								{
									m_ArrayDoc[i].Alcada = FALSE;
								}
							}
						}
						break;
					}
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
    /*********************************************
    ' * Verificando a Qtde de Contas no Malote   *
    ' * A Serem Vinculadas                       *
    ' *******************************************/
    iQtdContas = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CO" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdContas++;
	}

	if( iQtdContas == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

	/* Comentario do Vagner:
	   Nesta segunda fase o algoritmo tenta vincular um cheque com
	   uma combinacao das contas, esta combinacao comeca com n, depois
	   tenta n-1, n-2, ...
	   Porem ele so testa combinacoes com documentos adjacentes, 
	   ele nao verifica todas as combinacoes possiveis.
	*/

	/* Comentario do Vagner:
	   Nesta fase, se, e somente se, o Cheque de Pagamento for 
	   do Unibanco pode-se vincula-lo a qualquer Cobranca/Arrecadacao.
	*/

    /*****************************************
    ' * Segunda Fase do Vinculo              *
    ' * Vinculando Um Cheque a várias Contas *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto == 4 || Doc.TipoDocto == 5 || Doc.TipoDocto == 41) && 
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /*********************************************
            ' * Cheque, LI ou ADCC a Verificar o Vinculo *
            ' ********************************************/
			iInicio    = 1;
			iDesprezar = 1;
			while( iInicio <= iQtdContas )
			{
				ValorVinculo = 0;
				iConta       = 0;

				aIndConta.RemoveAll();
				for( j = 0; j < iQtdContas; j++ )
					aIndConta.Add(-1);

				for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux = m_ArrayDoc[j];
					if( DocAux.TipoGenerico == "CO" && 
						DocAux.Vinculo == 0 && 
						!DocAux.DesprezarVinculo )
					{
						iConta++;
						if( iConta == iInicio || iConta > iDesprezar )
						{
							ValorVinculo += DocAux.Valor;
							aIndConta[iConta-1] = j;
							if( ValorVinculo >= Doc.Valor )
								break;
						}
					}
				}
				if( iInicio % 50 == 0 )
				{
					PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
				}

				/**********************************************
				 * Verifica se o Valor da Combinacao bate com *
				 * o valor do cheque, entao vincula           *
				 **********************************************/
				if( ValorVinculo == Doc.Valor )
				{
					m_ArrayDoc[i].Vinculo = Doc.IdDocto;
					/**************************************
					' * Corrige o tipodocumento do cheque *
					' *************************************/
					if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
					{
						if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230") )
						{
							m_ArrayDoc[i].TipoDocto = 5;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
						else
						{
							m_ArrayDoc[i].TipoDocto = 6;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
					}

					for( j = 0; j < iQtdContas; j++ )
					{
						if( aIndConta[j] >= 0 )
						{
							m_ArrayDoc[aIndConta[j]].Vinculo = Doc.IdDocto;
						}
					}
					break;
				}

				iDesprezar++;
				if( iDesprezar >= iQtdContas )
				{
					iInicio++;
					iDesprezar = iInicio;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
    /********************************************
    ' * Verificando a Qtde de Cheques no Malote *
    ' * A Serem Vinculados                      *
    ' *******************************************/
    iQtdCheques = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto >= 4 && Doc.TipoDocto < 7 || Doc.TipoDocto == 41) &&
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdCheques++;
	}

	if( iQtdCheques == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

	/* Comentario do Vagner:
	   Nesta terceira fase o algoritmo tenta vincular uma conta com
	   uma combinacao dos cheques, esta combinacao comeca com n, depois
	   tenta n-1, n-2, ...
	   Porem ele so testa combinacoes com documentos adjacentes, 
	   ele nao verifica todas as combinacoes possiveis.
	*/

	/* Comentario do Vagner:
	   Nesta fase, se, e somente se, todos os Cheques de Pagamento forem 
	   do Unibanco pode-se vincula-los a qualquer Cobranca/Arrecadacao.
	*/

    /*****************************************
    ' * Terceira Fase do Vinculo              *
    ' * Vinculando Uma Conta a vários Cheques *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CO" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /********************************
            ' * Conta a Verificar o Vinculo *
            ' *******************************/
			iInicio    = 1;
			iDesprezar = 1;
			while( iInicio <= iQtdCheques )
			{
				ValorVinculo = 0;
				iCheque      = 0;
				iQtdChTerc	 = 0;
				iQtdChPagto  = 0;
				iQtdADCC     = 0;
				iQtdLI       = 0;

				aIndCheque.RemoveAll();
				for( j = 0; j < iQtdCheques; j++ )
					aIndCheque.Add(-1);

				for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux = m_ArrayDoc[j];
					/*
					if( (DocAux.TipoDocto == 4 || DocAux.TipoDocto == 5 || 
						 DocAux.TipoDocto == 41 ) && DocAux.Vinculo == 0 && 
						!DocAux.DesprezarVinculo )
					*/
					//Aceitar cheque 3o.
					if( DocAux.Vinculo == 0 && !DocAux.DesprezarVinculo &&
						(DocAux.TipoDocto == 4 || DocAux.TipoDocto == 5 ||
						 DocAux.TipoDocto == 6 || DocAux.TipoDocto == 41) )
					{
						iCheque++;
						if( iCheque == iInicio || iCheque > iDesprezar )
						{
							if( DocAux.TipoDocto == 4 )
								iQtdADCC++;
							else if( DocAux.TipoDocto == 5 )
								iQtdChPagto++;
							else if( DocAux.TipoDocto == 6 )
								iQtdChTerc++;
							else if( DocAux.TipoDocto == 41 )
								iQtdLI++;
							ValorVinculo += DocAux.Valor;
							aIndCheque[iCheque-1] = j;
							if( ValorVinculo >= Doc.Valor )
								break;
						}
					}
				}
				if( iInicio % 50 == 0 )
				{
					PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
				}

				/**********************************************
				 * Verifica se o Valor da Combinacao bate com *
				 * o valor da conta, entao vincula            *
				 **********************************************/
				if( ValorVinculo == Doc.Valor )
				{
					/* Nao pode haver cheque vinculado com ADCC  ou Cheque 3o. com LI na capa */
					if( iQtdChPagto == 0 || iQtdADCC == 0 &&
						(( iQtdChTerc > 0 && iQtdLI > 0 ) || iQtdChTerc == 0) )
					{
						m_ArrayDoc[i].Vinculo = Doc.IdDocto;
						for( j = 0; j < iQtdCheques; j++ )
						{
							if( aIndCheque[j] >= 0 )
							{
								m_ArrayDoc[aIndCheque[j]].Vinculo = Doc.IdDocto;
							}
						}
						break;
					}
				}

				iDesprezar++;
				if( iDesprezar >= iQtdCheques )
				{
					iInicio++;
					iDesprezar = iInicio;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
    /**********************************************
    ' * Verificando a Qtde de Cheques no Malote   *
    ' * A Serem Vinculados                        *
    ' ********************************************/
    iQtdCheques = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto >= 4 && Doc.TipoDocto < 7 || Doc.TipoDocto == 41) &&
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdCheques++;
	}

	if( iQtdCheques == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

	/* Comentario do Vagner:
	   Nesta parte, o processo vai tentar vincular um cheque para uma
	   conta mesmo que os valores nao batam, desde que a diferenca seja
	   menor ou igual a ValorAjusteContabil, e caso seja possivel vincular,
	   sera gerado o ajuste contabil.
	*/

    /*****************************************
    ' * Quarta Fase do Vinculo               *
    ' * Vinculando Um Cheque para Uma Conta  *
	' * com diferenca <= ValorAjusteContabil *
    ' ****************************************/

	/* Comentario do Vagner:
	   Nesta fase, se o Cheque de Pagamento for do Unibanco
	   pode-se vincula-lo a qualquer Cobranca/Arrecadacao.
	   Se o Cheque de Pagamento nao for do Unibanco
	   somente vincula-o se a Cobranca for diferente
	   de "Titulos Terceiros Sem CB" (12) e diferente de
	   "Cobranca Terceiros" (31)
	*/

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CP" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo &&
			Doc.TipoDocto != 41 ) /* Nao fazer ajuste contabil para LI */
		{
            /*********************************************
            ' * Cheque ou ADCC a Verificar o Vinculo     *
            ' ********************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.TipoGenerico == "CO" && 
					DocAux.Vinculo == 0 && 
					!DocAux.DesprezarVinculo &&
					abs(Doc.Valor - DocAux.Valor) > 0 &&
				    abs(Doc.Valor - DocAux.Valor) <= Converte(m_pParametro->m_ValorAjusteContabil) )
				{
					/************************************
					  Versao 2.7
					  Considerar Titulos do Bandeirantes 
					  olhando para o codigo de barras
					************************************/
					if( (Doc.TipoDocto == 6 && DocAux.TipoDocto != 12 && DocAux.TipoDocto != 31) || 
						 Doc.TipoDocto == 5 ||
						 Doc.TipoDocto == 4 ||
						 Doc.TipoDocto == 41 ||
						(DocAux.TipoDocto == 31 && DocAux.Leitura.Left(3) == _T("230")))
					{
						/**********************************
						' * Vinculando Conta com o Cheque *
						' *********************************/
						m_ArrayDoc[i].Vinculo  = Doc.IdDocto;
						m_ArrayDoc[j].Vinculo  = Doc.IdDocto;
						/**************************************
						' * Corrige o tipodocumento do cheque *
						' *************************************/
						if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
						{
							if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230") )
							{
								m_ArrayDoc[i].TipoDocto = 5;
								if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
								{
									m_ArrayDoc[i].Alcada = TRUE;
								}
								else
								{
									m_ArrayDoc[i].Alcada = FALSE;
								}
							}
							else
							{
								m_ArrayDoc[i].TipoDocto = 6;
								if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
								{
									m_ArrayDoc[i].Alcada = TRUE;
								}
								else
								{
									m_ArrayDoc[i].Alcada = FALSE;
								}
							}
						}

						/*****************************************
						' * Gravando Ajuste de Débito ou Crédito *
						' ****************************************/
						if( !InsereAjuste( (Doc.Valor > DocAux.Valor ? 42 : 43), 
										   0, "0", abs(Doc.Valor - DocAux.Valor), IdDocto ) )
						{
							return 0;
						}
						Ajuste.Alcada = FALSE;
						Ajuste.DesprezarVinculo = FALSE;
						Ajuste.IdDocto = IdDocto;
						Ajuste.TipoDocto = (Doc.Valor > DocAux.Valor ? 42 : 43);
						Ajuste.TipoGenerico = (Doc.Valor > DocAux.Valor ? "AC" : "AD");
						Ajuste.Valor = abs(Doc.Valor - DocAux.Valor);
						Ajuste.Vinculo = Doc.IdDocto;
						m_ArrayDoc.InsertAt(0, Ajuste, 1);

						break;
					}
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}
    
	
	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
	/**********************************
    ' * Verificar se ainda existe     *
    ' * Documentos a serem vinculados *
    ' *********************************/
    iQtdSemVinculo = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoGenerico == "CO" || Doc.TipoGenerico == "CP") &&
			 Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdSemVinculo++;
	}

	if( iQtdSemVinculo == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

    iQtdChPagto = 0;
	iQtdChTerc  = 0;
	iQtdADCC    = 0;
	iQtdLI      = 0;
	
	/******************************************
    ' * Verificar se pode Vincular o Restante *
    ' *****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoDocto == 5 )
		{
			iQtdChPagto++;
		}
		else if( Doc.TipoDocto >= 6 && Doc.TipoDocto <= 7 && Doc.Vinculo == 0 )
		{
			iQtdChTerc++;
		}
		else if( Doc.TipoDocto == 4 && Doc.Vinculo == 0 )
		{
			iQtdADCC++;
		}
		else if( Doc.TipoDocto == 41 && Doc.Vinculo == 0 )
		{
			iQtdLI++;
		}
		else if( Doc.DesprezarVinculo && Doc.Vinculo == 0 )
		{
			/* Comentario do Vagner:
			   Caso exista algum docto marcado para desprezar
			   vinculo, eh porque existem varios doctos com
			   mesmo valor, por isso nao se pode vincula-los
			   automaticamente
			*/
			return 1;
		}
	}

	// Acerto (Fase II): aceitar cheque de 3o. se tiver 1 ou + LI
	if( iQtdChTerc > 0 && iQtdLI == 0 )
	{
		 	
		return 1;

	}
	//Nao se pode vincular ADCC com Cheque	
	else if( iQtdChPagto > 0 && iQtdADCC > 0 )
	{

		return 1;

	}
	
	/* Comentario do Vagner:
	   Se o processo chegou ate este ponto, entao tudo que poderia
	   ser vinculado 1 para 1 ou 1 para N ja foi vinculado,
	   alem disso nao existem valores repetidos que poderiam gerar
	   varias alternativas de vinculo.
	   Entao, basta vincular todos os cheques de pagamento do Unibanco
	   que sobraram com as varias contas que sobraram, 
	   mas para isso os valores devem bater
	*/

    /**********************************
    ' * Verificar se Existe Diferença *
    ' *********************************/
	Diferenca    = 0;
	ValorCheques = 0;
	ValorContas  = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		//if( Doc.Vinculo == 0 && Doc.TipoDocto == 5 ) 
		// Acerto (Fase II): Aceitar LI e cheque 3o.
		if( Doc.Vinculo == 0 && (Doc.TipoDocto == 5 || Doc.TipoDocto == 6 || Doc.TipoDocto == 41))
		{
			ValorCheques += Doc.Valor;
		}
		else if( Doc.Vinculo == 0 && Doc.TipoGenerico == "CO" ) 
		{
			ValorContas += Doc.Valor;
		}
	}
	
	Diferenca = ValorCheques - ValorContas;

	/******************************************
	' * Quinta Fase do Vinculo                *
	' * Vinculando n Contas com n Cheques     *
	' *****************************************/
	if( Diferenca == 0 )
	{
		iVinculo = 0;

		for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
		{
			Doc = m_ArrayDoc[i];
			// Acerto (Fase II): Aceitar LI e cheque 3o.
			//if((Doc.TipoDocto == 5 || Doc.TipoGenerico == "CO") && Doc.Vinculo == 0 )
			if(Doc.Vinculo == 0 && (Doc.TipoGenerico == "CO" ||
				(Doc.TipoDocto == 5 || Doc.TipoDocto == 6 || Doc.TipoDocto == 41)))
			{
				/********************************
				' * Vinculando Contas e Cheques *
				' *******************************/
				if( iVinculo == 0 )
					iVinculo = Doc.IdDocto;

				m_ArrayDoc[i].Vinculo = iVinculo;
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE; 
		ContemCOTerceiro = TRUE;
	}
    /*********************************************
    ' * Verificando a Qtde de Contas no Malote   *
    ' * A Serem Vinculadas                       *
    ' *******************************************/
    iQtdContas = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CO" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdContas++;
	}

	if( iQtdContas == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

	/* Comentario do Vagner:
	   Nesta fase, se, e somente se, o Cheque de Pagamento for 
	   do Unibanco pode-se vincula-lo a qualquer Cobranca/Arrecadacao.
	*/

    /*****************************************
    ' * Sexta Fase do Vinculo                *
    ' * Vinculando Um Cheque a várias Contas *
	' * com diferenca <= ValorAjusteContabil *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto == 4 || Doc.TipoDocto == 5 ) && 
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo ) 
			/* Nao fazer ajuste contabil para LI */
		{
            /*********************************************
            ' * Cheque ou ADCC a Verificar o Vinculo     *
            ' ********************************************/
			iInicio    = 1;
			iDesprezar = 1;
			while( iInicio <= iQtdContas )
			{
				ValorVinculo = 0;
				iConta       = 0;

				aIndConta.RemoveAll();
				for( j = 0; j < iQtdContas; j++ )
					aIndConta.Add(-1);

				for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux = m_ArrayDoc[j];
					if( DocAux.TipoGenerico == "CO" && 
						DocAux.Vinculo == 0 && 
						!DocAux.DesprezarVinculo )
					{
						iConta++;
						if( iConta == iInicio || iConta > iDesprezar )
						{
							ValorVinculo += DocAux.Valor;
							aIndConta[iConta-1] = j;
							if( abs(Doc.Valor - ValorVinculo) <= 
								Converte(m_pParametro->m_ValorAjusteContabil) )
								break;
						}
					}
				}
				if( iInicio % 50 == 0 )
				{
					PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
				}

				/*******************************************************
				 * Verifica se o Valor da Combinacao menos             *
				 * o valor do cheque, e menor que ValorAjusteContabil  *
				 *******************************************************/
				if( abs(Doc.Valor - ValorVinculo) > 0 &&
					abs(Doc.Valor - ValorVinculo) <= 
					Converte(m_pParametro->m_ValorAjusteContabil) )
				{
					m_ArrayDoc[i].Vinculo = Doc.IdDocto;
					/**************************************
					' * Corrige o tipodocumento do cheque *
					' *************************************/
					if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
					{
						if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230") )
						{
							m_ArrayDoc[i].TipoDocto = 5;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
						else
						{
							m_ArrayDoc[i].TipoDocto = 6;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
					}

					for( j = 0; j < iQtdContas; j++ )
					{
						if( aIndConta[j] >= 0 )
						{
							m_ArrayDoc[aIndConta[j]].Vinculo = Doc.IdDocto;
						}
					}

					/*****************************************
					' * Gravando Ajuste de Débito ou Crédito *
					' ****************************************/
					if( !InsereAjuste( (Doc.Valor > ValorVinculo ? 42 : 43), 
									   0, "0", abs(Doc.Valor - ValorVinculo), IdDocto ) )
					{
						return 0;
					}
					Ajuste.Alcada = FALSE;
					Ajuste.DesprezarVinculo = FALSE;
					Ajuste.IdDocto = IdDocto;
					Ajuste.TipoDocto = (Doc.Valor > ValorVinculo ? 42 : 43);
					Ajuste.TipoGenerico = (Doc.Valor > ValorVinculo ? "AC" : "AD");
					Ajuste.Valor = abs(Doc.Valor - ValorVinculo);
					Ajuste.Vinculo = Doc.IdDocto;
					m_ArrayDoc.InsertAt(0, Ajuste, 1);
					
					break;
				}

				iDesprezar++;
				if( iDesprezar >= iQtdContas )
				{
					iInicio++;
					iDesprezar = iInicio;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
    /********************************************
    ' * Verificando a Qtde de Cheques no Malote *
    ' * A Serem Vinculados                      *
    ' *******************************************/
    iQtdCheques = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto >= 4 && Doc.TipoDocto < 7 || Doc.TipoDocto == 41) &&
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdCheques++;
	}

	if( iQtdCheques == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

	/* Comentario do Vagner:
	   Nesta fase, se, e somente se, todos os Cheques de Pagamento forem 
	   do Unibanco pode-se vincula-los a qualquer Cobranca/Arrecadacao.
	*/

    /******************************************
    ' * Setima Fase do Vinculo                *
    ' * Vinculando Uma Conta a vários Cheques *
	' * com diferenca <= ValorAjusteContabil  *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CO" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /********************************
            ' * Conta a Verificar o Vinculo *
            ' *******************************/
			iInicio    = 1;
			iDesprezar = 1;
			while( iInicio <= iQtdCheques )
			{
				ValorVinculo = 0;
				iCheque      = 0;
				iQtdChPagto  = 0;
				iQtdADCC     = 0;

				aIndCheque.RemoveAll();
				for( j = 0; j < iQtdCheques; j++ )
					aIndCheque.Add(-1);

				for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux = m_ArrayDoc[j];
					if( (DocAux.TipoDocto == 4 || DocAux.TipoDocto == 5) && 
						DocAux.Vinculo == 0 && !DocAux.DesprezarVinculo )
						/* Nao fazer ajuste contabil para LI */
					{
						iCheque++;
						if( iCheque == iInicio || iCheque > iDesprezar )
						{
							if( DocAux.TipoDocto == 4 )
								iQtdADCC++;
							else if( DocAux.TipoDocto == 5 )
								iQtdChPagto++;
							ValorVinculo += DocAux.Valor;
							aIndCheque[iCheque-1] = j;
							if( abs(ValorVinculo - Doc.Valor) <= 
								Converte(m_pParametro->m_ValorAjusteContabil) )
								break;
						}
					}
				}
				if( iInicio % 50 == 0 )
				{
					PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
				}

				/*****************************************************
				 * Verifica se o Valor da Combinacao menos           *
				 * o valor da conta, e menor que ValorAjusteContabil *
				 *****************************************************/
				if( abs(ValorVinculo - Doc.Valor) > 0 &&
					abs(ValorVinculo - Doc.Valor) <= 
					Converte(m_pParametro->m_ValorAjusteContabil) )
				{
					/* Nao pode haver cheque vinculado com ADCC */
					if( iQtdChPagto == 0 || iQtdADCC == 0 )
					{
						m_ArrayDoc[i].Vinculo = Doc.IdDocto;
						for( j = 0; j < iQtdCheques; j++ )
						{
							if( aIndCheque[j] >= 0 )
							{
								m_ArrayDoc[aIndCheque[j]].Vinculo = Doc.IdDocto;
							}
						}

						/*****************************************
						' * Gravando Ajuste de Débito ou Crédito *
						' ****************************************/
						if( !InsereAjuste( (ValorVinculo > Doc.Valor ? 42 : 43), 
										   0, "0", abs(ValorVinculo - Doc.Valor), IdDocto ) )
						{
							return 0;
						}
						Ajuste.Alcada = FALSE;
						Ajuste.DesprezarVinculo = FALSE;
						Ajuste.IdDocto = IdDocto;
						Ajuste.TipoDocto = (ValorVinculo > Doc.Valor ? 42 : 43);
						Ajuste.TipoGenerico = (ValorVinculo > Doc.Valor ? "AC" : "AD");
						Ajuste.Valor = abs(ValorVinculo - Doc.Valor);
						Ajuste.Vinculo = Doc.IdDocto;
						m_ArrayDoc.InsertAt(0, Ajuste, 1);
						
						break;
					}
				}

				iDesprezar++;
				if( iDesprezar >= iQtdCheques )
				{
					iInicio++;
					iDesprezar = iInicio;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
	/**********************************
    ' * Verificar se ainda existe     *
    ' * Documentos a serem vinculados *
    ' *********************************/
    iQtdSemVinculo = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoGenerico == "CO" || Doc.TipoGenerico == "CP") &&
			 Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdSemVinculo++;
	}

	if( iQtdSemVinculo == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

    /**********************************
    ' * Verificar se Existe Diferença *
    ' *********************************/
	Diferenca    = 0;
	ValorCheques = 0;
	ValorContas  = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.Vinculo == 0 && Doc.TipoDocto == 5 )  /* Nao fazer ajuste contabil para LI */
		{
			ValorCheques += Doc.Valor;
		}
		else if( Doc.Vinculo == 0 && Doc.TipoGenerico == "CO" ) 
		{
			ValorContas += Doc.Valor;
		}
	}
	
	Diferenca = ValorCheques - ValorContas;

	/******************************************
	' * Oitava Fase do Vinculo                *
	' * Vinculando n Contas com n Cheques     *
	' * com diferenca <= ValorAjusteContabil  *
	' *****************************************/
	if( abs(Diferenca) >= 0 && 
		abs(Diferenca) <= Converte(m_pParametro->m_ValorAjusteContabil) )
	{
		iVinculo = 0;

		for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
		{
			Doc = m_ArrayDoc[i];
			if( (Doc.TipoDocto == 5 || Doc.TipoGenerico == "CO") 
				 && Doc.Vinculo == 0 )
			{
				/********************************
				' * Vinculando Contas e Cheques *
				' *******************************/
				if( iVinculo == 0 )
					iVinculo = Doc.IdDocto;

				m_ArrayDoc[i].Vinculo = iVinculo;
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}

		if( abs(Diferenca) > 0 )
		{
			/*****************************************
			' * Gravando Ajuste de Débito ou Crédito *
			' ****************************************/
			if( !InsereAjuste( (Diferenca > 0 ? 42 : 43), 
							   0, "0", abs(Diferenca), IdDocto ) )
			{
				return 0;
			}
			Ajuste.Alcada = FALSE;
			Ajuste.DesprezarVinculo = FALSE;
			Ajuste.IdDocto = IdDocto;
			Ajuste.TipoDocto = (Diferenca > 0 ? 42 : 43);
			Ajuste.TipoGenerico = (Diferenca > 0 ? "AC" : "AD");
			Ajuste.Valor = abs(Diferenca);
			Ajuste.Vinculo = iVinculo;
			m_ArrayDoc.InsertAt(0, Ajuste, 1);
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	if( lVinculo )
	{
		RemoveVinculo( lVinculo );
		RemoveAjustes();
	}

	// Sucesso !

	return 1;

}

// retorna: -1 se erro; 0 se nao vinculou; 1 se vinculou
int CVinculador::VinculaDocumentoMaloteRegraAntiga( void )
{
	DOCUMENTO Doc, DocAux, Ajuste;
	int i, j, iQtdCheques, iQtdContas;
	int iQtdSemVinculo, iQtdADCC, iQtdLI, iQtdChPagto;
	long iVinculo, lVinculo;
	__int64 ValorVinculo, Diferenca, ValorCheques, ValorContas;
	int iConta, iCheque;
	int iInicio, iDesprezar;
	long IdDocto;
	BOOL ContemCPTerceiro, ContemCOTerceiro;

	CArray<int, int> aIndConta;
	CArray<int, int> aIndCheque;

	lVinculo = 0;

	/* Comentario do Vagner:
	   Este algoritmo foi desenvolvido pela Proservvi,
	   por absoluta falta de tempo, ele nao pode ser
	   melhorado, mas ele funciona na grande maioria dos casos.
	*/

	/*****************************************
    ' * Marcar para Desprezar para o Vínculo *
    ' * Os Cheques com o mesmo Valor         *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		iQtdCheques = 0;
		iQtdContas  = 0;

		Doc = m_ArrayDoc[i];
		
		if( Doc.DesprezarVinculo == FALSE && Doc.Vinculo == 0 && 
			Doc.TipoDocto > 3 && Doc.TipoDocto <= 6 )
		{
            /****************************
            ' * Cheque/ADCC sem Vínculo *
            ' ***************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.Vinculo == 0 && DocAux.Valor == Doc.Valor )
				{
					if( DocAux.TipoDocto > 3 && DocAux.TipoDocto <= 6 )
					{
                        /*******************************
                        ' * Cheque/ADCC no Mesmo Valor *
                        ' ******************************/
						iQtdCheques++;
					}
					else if( DocAux.TipoGenerico == "CO" )
					{
                        /*************************
                        ' * Conta no mesmo Valor *
                        ' ************************/
						iQtdContas++;
					}
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}

		if( iQtdContas > 1 && m_QtdChequePagto > 1 )
		{
            /*************************************************
            ' * Desprezar para o Vínculo Automático          *
            ' ************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdContas > 1 && iQtdCheques > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdCheques > 1 && m_QtdContas > 1 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdCheques > 1 && iQtdContas > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}

		if( m_ArrayDoc[i].DesprezarVinculo )
		{
            /*****************************************
            ' * Marcar para Desprezar para o Vínculo *
            ' * Todos que tenham o mesmo valor       *
            ' ****************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				if( Doc.Valor == m_ArrayDoc[j].Valor )
					m_ArrayDoc[j].DesprezarVinculo = TRUE;
			}
		}
	}

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		iQtdContas     = 0;
		iQtdCheques    = 0;

		Doc = m_ArrayDoc[i];
		if( Doc.DesprezarVinculo == FALSE && Doc.Vinculo == 0 && 
			Doc.TipoGenerico == "CO" )
		{
            /***********************
            ' * Cheque sem Vínculo *
            ' **********************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.Vinculo == 0 && DocAux.Valor == Doc.Valor )
				{
					if( DocAux.TipoDocto > 3 && DocAux.TipoDocto <= 6 )
						iQtdCheques++;
					else if( DocAux.TipoGenerico == "CO" )
						iQtdContas++;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}


		if( iQtdCheques > 1 && m_QtdContas > 1 )
		{
            /*************************************************
            ' * Desprezar para o Vínculo Automático          *
            ' ************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdCheques > 1 && iQtdContas > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdContas > 1 && m_QtdChequePagto > 1 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdContas > 1 && iQtdCheques > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}

		if( m_ArrayDoc[i].DesprezarVinculo )
		{
            /*****************************************
            ' * Marcar para Desprezar para o Vínculo *
            ' * Todos que tenham o mesmo valor       *
            ' ****************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				if( Doc.Valor == m_ArrayDoc[j].Valor )
					m_ArrayDoc[j].DesprezarVinculo = TRUE;
			}
		}
		
	}


    /****************************************
    ' * Primeira Fase do Vinculo            *
    ' * Vinculando Um Cheque para Uma Conta *
    ' ***************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CP" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /*****************************************
            ' * Cheque ou ADCC a Verificar o Vinculo *
            ' ****************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.TipoGenerico == "CO" && 
					DocAux.Vinculo == 0 && 
					!DocAux.DesprezarVinculo &&
					Doc.Valor == DocAux.Valor )
				{
                    /**********************************
                    ' * Vinculando Conta com o Cheque *
                    ' *********************************/
					m_ArrayDoc[i].Vinculo  = Doc.IdDocto;
					m_ArrayDoc[j].Vinculo  = Doc.IdDocto;
					/**************************************
					' * Corrige o tipodocumento do cheque *
					' *************************************/
					if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
					{
						if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230"))
						{
							m_ArrayDoc[i].TipoDocto = 5;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
						else
						{
							m_ArrayDoc[i].TipoDocto = 6;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
					}
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
    /*******************************************
    ' * Verificando a Qtde de Contas no Malote *
    ' * A Serem Vinculadas                     *
    ' ******************************************/
    iQtdContas = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CO" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdContas++;
	}

	if( iQtdContas == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

	/* Comentario do Vagner:
	   Nesta segunda fase o algoritmo tenta vincular um cheque com
	   uma combinacao das contas, esta combinacao comeca com n, depois
	   tenta n-1, n-2, ...
	   Porem ele so testa combinacoes com documentos adjacentes, 
	   ele nao verifica todas as combinacoes possiveis.
	*/

    /*****************************************
    ' * Segunda Fase do Vinculo              *
    ' * Vinculando Um Cheque a várias Contas *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CP" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /*****************************************
            ' * Cheque ou ADCC a Verificar o Vinculo *
            ' ****************************************/
			iInicio    = 1;
			iDesprezar = 1;
			while( iInicio <= iQtdContas )
			{
				ValorVinculo = 0;
				iConta       = 0;

				aIndConta.RemoveAll();
				for( j = 0; j < iQtdContas; j++ )
					aIndConta.Add(-1);

				for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux = m_ArrayDoc[j];
					if( DocAux.TipoGenerico == "CO" && 
						DocAux.Vinculo == 0 && 
						!DocAux.DesprezarVinculo )
					{
						iConta++;
						if( iConta == iInicio || iConta > iDesprezar )
						{
							ValorVinculo += DocAux.Valor;
							aIndConta[iConta-1] = j;
							if( ValorVinculo >= Doc.Valor )
								break;
						}
					}
				}
				if( iInicio % 50 == 0 )
				{
					PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
				}

				/**********************************************
				 * Verifica se o Valor da Combinacao bate com *
				 * o valor do cheque, entao vincula           *
				 **********************************************/
				if( ValorVinculo == Doc.Valor )
				{
					m_ArrayDoc[i].Vinculo = Doc.IdDocto;
					/**************************************
					' * Corrige o tipodocumento do cheque *
					' *************************************/
					if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
					{
						if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230") )
						{
							m_ArrayDoc[i].TipoDocto = 5;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
						else
						{
							m_ArrayDoc[i].TipoDocto = 6;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
					}

					for( j = 0; j < iQtdContas; j++ )
					{
						if( aIndConta[j] >= 0 )
						{
							m_ArrayDoc[aIndConta[j]].Vinculo = Doc.IdDocto;
						}
					}
					break;
				}

				iDesprezar++;
				if( iDesprezar >= iQtdContas )
				{
					iInicio++;
					iDesprezar = iInicio;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
    /********************************************
    ' * Verificando a Qtde de Cheques no Malote *
    ' * A Serem Vinculados                      *
    ' *******************************************/
    iQtdCheques = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto >= 4 && Doc.TipoDocto < 7 || Doc.TipoDocto == 41) &&
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdCheques++;
	}

	if( iQtdCheques == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

	/* Comentario do Vagner:
	   Nesta segunda fase o algoritmo tenta vincular uma conta com
	   uma combinacao dos cheques, esta combinacao comeca com n, depois
	   tenta n-1, n-2, ...
	   Porem ele so testa combinacoes com documentos adjacentes, 
	   ele nao verifica todas as combinacoes possiveis.
	*/

    /******************************************
    ' * Terceira Fase do Vinculo              *
    ' * Vinculando Uma Conta a vários Cheques *
    ' *****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CO" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /********************************
            ' * Conta a Verificar o Vinculo *
            ' *******************************/
			iInicio    = 1;
			iDesprezar = 1;
			while( iInicio <= iQtdCheques )
			{
				ValorVinculo = 0;
				iCheque      = 0;
				iQtdChPagto  = 0;
				iQtdADCC     = 0;
				iQtdLI       = 0;

				aIndCheque.RemoveAll();
				for( j = 0; j < iQtdCheques; j++ )
					aIndCheque.Add(-1);

				for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux = m_ArrayDoc[j];
					if( DocAux.TipoGenerico == "CP" && 
						DocAux.Vinculo == 0 && 
						!DocAux.DesprezarVinculo )
					{
						iCheque++;
						if( iCheque == iInicio || iCheque > iDesprezar )
						{
							if( DocAux.TipoDocto == 4 )
								iQtdADCC++;
							else if( DocAux.TipoDocto >= 5 && DocAux.TipoDocto <= 7 )
								iQtdChPagto++;
							else if( DocAux.TipoDocto == 41 )
								iQtdLI++;
							ValorVinculo += DocAux.Valor;
							aIndCheque[iCheque-1] = j;
							if( ValorVinculo >= Doc.Valor )
								break;
						}
					}
				}
				if( iInicio % 50 == 0 )
				{
					PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
				}

				/**********************************************
				 * Verifica se o Valor da Combinacao bate com *
				 * o valor da conta, entao vincula            *
				 **********************************************/
				if( ValorVinculo == Doc.Valor )
				{
					/* Nao pode haver cheque vinculado com ADCC */
					if( iQtdChPagto == 0 || iQtdADCC == 0 )
					{
						m_ArrayDoc[i].Vinculo = Doc.IdDocto;
						for( j = 0; j < iQtdCheques; j++ )
						{
							if( aIndCheque[j] >= 0 )
							{
								m_ArrayDoc[aIndCheque[j]].Vinculo = Doc.IdDocto;
								/***********************************************
								 * Se cheque Ubb transforma em cheque terceiro *
								 ***********************************************/
								if( m_ArrayDoc[aIndCheque[j]].TipoDocto == 5  ||
									m_ArrayDoc[aIndCheque[j]].TipoDocto == 7 )
								{
									m_ArrayDoc[aIndCheque[j]].TipoDocto = 6;
									if(	m_ArrayDoc[aIndCheque[j]].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
									{
										m_ArrayDoc[aIndCheque[j]].Alcada = TRUE;
									}
									else
									{
										m_ArrayDoc[aIndCheque[j]].Alcada = FALSE;
									}
								}
							}
						}
						break;
					}
				}

				iDesprezar++;
				if( iDesprezar >= iQtdCheques )
				{
					iInicio++;
					iDesprezar = iInicio;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
    /**********************************************
    ' * Verificando a Qtde de Cheques no Malote   *
    ' * A Serem Vinculados                        *
    ' ********************************************/
    iQtdCheques = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto >= 4 && Doc.TipoDocto < 7 || Doc.TipoDocto == 41) &&
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdCheques++;
	}

	if( iQtdCheques == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

	/* Comentario do Vagner:
	   Nesta parte, o processo vai tentar vincular um cheque para uma
	   conta mesmo que os valores nao batam, desde que a diferenca seja
	   menor ou igual a ValorAjusteContabil, e caso seja possivel vincular,
	   sera gerado o ajuste contabil.
	*/

    /*****************************************
    ' * Quarta Fase do Vinculo               *
    ' * Vinculando Um Cheque para Uma Conta  *
	' * com diferenca <= ValorAjusteContabil *
    ' ****************************************/

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CP" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo &&
			Doc.TipoDocto != 41 ) /* Nao fazer ajuste contabil para LI */
		{
            /*********************************************
            ' * Cheque ou ADCC a Verificar o Vinculo     *
            ' ********************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.TipoGenerico == "CO" && 
					DocAux.Vinculo == 0 && 
					!DocAux.DesprezarVinculo &&
					abs(Doc.Valor - DocAux.Valor) > 0 &&
				    abs(Doc.Valor - DocAux.Valor) <= Converte(m_pParametro->m_ValorAjusteContabil) )
				{
					/**********************************
					' * Vinculando Conta com o Cheque *
					' *********************************/
					m_ArrayDoc[i].Vinculo  = Doc.IdDocto;
					m_ArrayDoc[j].Vinculo  = Doc.IdDocto;
					/**************************************
					' * Corrige o tipodocumento do cheque *
					' *************************************/
					if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
					{
						if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230") )
						{
							m_ArrayDoc[i].TipoDocto = 5;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
						else
						{
							m_ArrayDoc[i].TipoDocto = 6;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
					}

					/*****************************************
					' * Gravando Ajuste de Débito ou Crédito *
					' ****************************************/
					if( !InsereAjuste( (Doc.Valor > DocAux.Valor ? 42 : 43), 
									   0, "0", abs(Doc.Valor - DocAux.Valor), IdDocto ) )
					{
						return 0;
					}
					Ajuste.Alcada = FALSE;
					Ajuste.DesprezarVinculo = FALSE;
					Ajuste.IdDocto = IdDocto;
					Ajuste.TipoDocto = (Doc.Valor > DocAux.Valor ? 42 : 43);
					Ajuste.TipoGenerico = (Doc.Valor > DocAux.Valor ? "AC" : "AD");
					Ajuste.Valor = abs(Doc.Valor - DocAux.Valor);
					Ajuste.Vinculo = Doc.IdDocto;
					m_ArrayDoc.InsertAt(0, Ajuste, 1);

					break;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
	/**********************************
    ' * Verificar se ainda existe     *
    ' * Documentos a serem vinculados *
    ' *********************************/
    iQtdSemVinculo = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoGenerico == "CO" || Doc.TipoGenerico == "CP") &&
			 Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdSemVinculo++;
	}

	if( iQtdSemVinculo == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

    iQtdChPagto = 0;
	iQtdADCC    = 0;
	iQtdLI      = 0;
	
	/******************************************
    ' * Verificar se pode Vincular o Restante *
    ' *****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoDocto >= 5 && Doc.TipoDocto <= 7 && Doc.Vinculo == 0 )
		{
			iQtdChPagto++;
		}
		else if( Doc.TipoDocto == 4 && Doc.Vinculo == 0 )
		{
			iQtdADCC++;
		}
		else if( Doc.TipoDocto == 41 && Doc.Vinculo == 0 )
		{
			iQtdLI++;
		}
		else if( Doc.DesprezarVinculo && Doc.Vinculo == 0 )
		{
			/* Comentario do Vagner:
			   Caso exista algum docto marcado para desprezar
			   vinculo, eh porque existem varios doctos com
			   mesmo valor, por isso nao se pode vincula-los
			   automaticamente
			*/
			return 1;
		}
	}

	if( iQtdChPagto > 0 && iQtdADCC > 0 )
	{
		// Nao se pode vincular ADCC com Cheque
		return 1;

	}
	
	/* Comentario do Vagner:
	   Se o processo chegou ate este ponto, entao tudo que poderia
	   ser vinculado 1 para 1 ou 1 para N ja foi vinculado,
	   alem disso nao existem valores repetidos que poderiam gerar
	   varias alternativas de vinculo.
	   Entao, basta vincular todos os cheques de pagamento do Unibanco
	   que sobraram com as varias contas que sobraram, 
	   mas para isso os valores devem bater
	*/
	
	/**********************************
    ' * Verificar se Existe Diferença *
    ' *********************************/
	Diferenca    = 0;
	ValorCheques = 0;
	ValorContas  = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.Vinculo == 0 && (Doc.TipoGenerico == "CP" || Doc.TipoDocto == 38) ) // AD
		{
			ValorCheques += Doc.Valor;
		}
		else if( Doc.Vinculo == 0 && (Doc.TipoGenerico == "CO" || Doc.TipoDocto == 34) ) // AC
		{
			ValorContas += Doc.Valor;
		}
	}
	
	Diferenca = ValorCheques - ValorContas;
    
	/******************************************
	' * Quinta Fase do Vinculo                *
	' * Vinculando n Contas com n Cheques     *
	' *****************************************/
	if( Diferenca == 0 )
	{
		iVinculo = 0;

		for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
		{
			Doc = m_ArrayDoc[i];
			if( (Doc.TipoGenerico == "CP" || Doc.TipoGenerico == "CO") 
				 && Doc.Vinculo == 0 )
			{
				/********************************
				' * Vinculando Contas e Cheques *
				' *******************************/
				if( iVinculo == 0 )
					iVinculo = Doc.IdDocto;

				m_ArrayDoc[i].Vinculo = iVinculo;
				/***********************************************
				 * Se cheque Ubb transforma em cheque terceiro *
				 ***********************************************/
				if( Doc.TipoDocto == 5 || Doc.TipoDocto == 7)
				{
					m_ArrayDoc[i].TipoDocto = 6;
					if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
					{
						m_ArrayDoc[i].Alcada = TRUE;
					}
					else
					{
						m_ArrayDoc[i].Alcada = FALSE;
					}
				}

			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}
	
	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
    /*********************************************
    ' * Verificando a Qtde de Contas no Malote   *
    ' * A Serem Vinculadas                       *
    ' *******************************************/
    iQtdContas = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CO" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdContas++;
	}

	if( iQtdContas == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

    /*****************************************
    ' * Sexta Fase do Vinculo                *
    ' * Vinculando Um Cheque a várias Contas *
	' * com diferenca <= ValorAjusteContabil *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CP" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo &&
			Doc.TipoDocto != 41 ) /* Nao fazer ajuste contabil para LI */
		{
            /*********************************************
            ' * Cheque ou ADCC a Verificar o Vinculo     *
            ' ********************************************/
			iInicio    = 1;
			iDesprezar = 1;
			while( iInicio <= iQtdContas )
			{
				ValorVinculo = 0;
				iConta       = 0;

				aIndConta.RemoveAll();
				for( j = 0; j < iQtdContas; j++ )
					aIndConta.Add(-1);

				for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux = m_ArrayDoc[j];
					if( DocAux.TipoGenerico == "CO" && 
						DocAux.Vinculo == 0 && 
						!DocAux.DesprezarVinculo )
					{
						iConta++;
						if( iConta == iInicio || iConta > iDesprezar )
						{
							ValorVinculo += DocAux.Valor;
							aIndConta[iConta-1] = j;
							if( abs(Doc.Valor - ValorVinculo) <= 
								Converte(m_pParametro->m_ValorAjusteContabil) )
								break;
						}
					}
				}
				if( iInicio % 50 == 0 )
				{
					PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
				}

				/*******************************************************
				 * Verifica se o Valor da Combinacao menos             *
				 * o valor do cheque, e menor que ValorAjusteContabil  *
				 *******************************************************/
				if( abs(Doc.Valor - ValorVinculo) > 0 &&
					abs(Doc.Valor - ValorVinculo) <= 
					Converte(m_pParametro->m_ValorAjusteContabil) )
				{
					m_ArrayDoc[i].Vinculo = Doc.IdDocto;
					/**************************************
					' * Corrige o tipodocumento do cheque *
					' *************************************/
					if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
					{
						if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230") )
						{
							m_ArrayDoc[i].TipoDocto = 5;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
						else
						{
							m_ArrayDoc[i].TipoDocto = 6;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
					}

					for( j = 0; j < iQtdContas; j++ )
					{
						if( aIndConta[j] >= 0 )
						{
							m_ArrayDoc[aIndConta[j]].Vinculo = Doc.IdDocto;
						}
					}

					/*****************************************
					' * Gravando Ajuste de Débito ou Crédito *
					' ****************************************/
					if( !InsereAjuste( (Doc.Valor > ValorVinculo ? 42 : 43), 
									   0, "0", abs(Doc.Valor - ValorVinculo), IdDocto ) )
					{
						return 0;
					}
					Ajuste.Alcada = FALSE;
					Ajuste.DesprezarVinculo = FALSE;
					Ajuste.IdDocto = IdDocto;
					Ajuste.TipoDocto = (Doc.Valor > ValorVinculo ? 42 : 43);
					Ajuste.TipoGenerico = (Doc.Valor > ValorVinculo ? "AC" : "AD");
					Ajuste.Valor = abs(Doc.Valor - ValorVinculo);
					Ajuste.Vinculo = Doc.IdDocto;
					m_ArrayDoc.InsertAt(0, Ajuste, 1);
					
					break;
				}

				iDesprezar++;
				if( iDesprezar >= iQtdContas )
				{
					iInicio++;
					iDesprezar = iInicio;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
    /********************************************
    ' * Verificando a Qtde de Cheques no Malote *
    ' * A Serem Vinculados                      *
    ' *******************************************/
    iQtdCheques = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto >= 4 && Doc.TipoDocto < 7 || Doc.TipoDocto == 41) &&
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdCheques++;
	}

	if( iQtdCheques == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

    /******************************************
    ' * Setima Fase do Vinculo                *
    ' * Vinculando Uma Conta a vários Cheques *
	' * com diferenca <= ValorAjusteContabil  *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CO" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /********************************
            ' * Conta a Verificar o Vinculo *
            ' *******************************/
			iInicio    = 1;
			iDesprezar = 1;
			while( iInicio <= iQtdCheques )
			{
				ValorVinculo = 0;
				iCheque      = 0;
				iQtdChPagto  = 0;
				iQtdADCC     = 0;

				aIndCheque.RemoveAll();
				for( j = 0; j < iQtdCheques; j++ )
					aIndCheque.Add(-1);

				for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux = m_ArrayDoc[j];
					if( DocAux.TipoGenerico == "CP" && 
						DocAux.Vinculo == 0 && !DocAux.DesprezarVinculo &&
						DocAux.TipoDocto != 41 ) /* Nao fazer ajuste contabil para LI */
					{
						iCheque++;
						if( iCheque == iInicio || iCheque > iDesprezar )
						{
							if( DocAux.TipoDocto == 4 )
								iQtdADCC++;
							else if( DocAux.TipoDocto >= 5 && DocAux.TipoDocto <= 7 )
								iQtdChPagto++;
							ValorVinculo += DocAux.Valor;
							aIndCheque[iCheque-1] = j;
							if( abs(ValorVinculo - Doc.Valor) <= 
								Converte(m_pParametro->m_ValorAjusteContabil) )
								break;
						}
					}
				}
				if( iInicio % 50 == 0 )
				{
					PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
				}

				/*****************************************************
				 * Verifica se o Valor da Combinacao menos           *
				 * o valor da conta, e menor que ValorAjusteContabil *
				 *****************************************************/
				if( abs(ValorVinculo - Doc.Valor) > 0 &&
					abs(ValorVinculo - Doc.Valor) <= 
					Converte(m_pParametro->m_ValorAjusteContabil) )
				{
					/* Nao pode haver cheque vinculado com ADCC */
					if( iQtdChPagto == 0 || iQtdADCC == 0 )
					{
						m_ArrayDoc[i].Vinculo = Doc.IdDocto;
						for( j = 0; j < iQtdCheques; j++ )
						{
							if( aIndCheque[j] >= 0 )
							{
								m_ArrayDoc[aIndCheque[j]].Vinculo = Doc.IdDocto;
								/***********************************************
								 * Se cheque Ubb transforma em cheque terceiro *
								 ***********************************************/
								if( m_ArrayDoc[aIndCheque[j]].TipoDocto == 5  ||
									m_ArrayDoc[aIndCheque[j]].TipoDocto == 7 )
								{
									m_ArrayDoc[aIndCheque[j]].TipoDocto = 6;
									if(	m_ArrayDoc[aIndCheque[j]].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
									{
										m_ArrayDoc[aIndCheque[j]].Alcada = TRUE;
									}
									else
									{
										m_ArrayDoc[aIndCheque[j]].Alcada = FALSE;
									}
								}
							}
						}

						/*****************************************
						' * Gravando Ajuste de Débito ou Crédito *
						' ****************************************/
						if( !InsereAjuste( (ValorVinculo > Doc.Valor ? 42 : 43), 
										   0, "0", abs(ValorVinculo - Doc.Valor), IdDocto ) )
						{
							return 0;
						}
						Ajuste.Alcada = FALSE;
						Ajuste.DesprezarVinculo = FALSE;
						Ajuste.IdDocto = IdDocto;
						Ajuste.TipoDocto = (ValorVinculo > Doc.Valor ? 42 : 43);
						Ajuste.TipoGenerico = (ValorVinculo > Doc.Valor ? "AC" : "AD");
						Ajuste.Valor = abs(ValorVinculo - Doc.Valor);
						Ajuste.Vinculo = Doc.IdDocto;
						m_ArrayDoc.InsertAt(0, Ajuste, 1);
						
						break;
					}
				}

				iDesprezar++;
				if( iDesprezar >= iQtdCheques )
				{
					iInicio++;
					iDesprezar = iInicio;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	ContemCPTerceiro = FALSE;
	ContemCOTerceiro = FALSE;
	if( lVinculo )
	{
		ContemCPTerceiro = TRUE;
		ContemCOTerceiro = TRUE;
	}
	/**********************************
    ' * Verificar se ainda existe     *
    ' * Documentos a serem vinculados *
    ' *********************************/
    iQtdSemVinculo = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoGenerico == "CO" || Doc.TipoGenerico == "CP") &&
			 Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdSemVinculo++;
	}

	if( iQtdSemVinculo == 0 )
	{
		if( ContemCPTerceiro && ContemCOTerceiro )
		{
			RemoveVinculo( lVinculo );
			RemoveAjustes();
		}
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return 1;
	}

	/**********************************
    ' * Verificar se Existe Diferença *
    ' *********************************/
	Diferenca    = 0;
	ValorCheques = 0;
	ValorContas  = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.Vinculo == 0 && (Doc.TipoGenerico == "CP" || Doc.TipoDocto == 38) && 
			Doc.TipoDocto != 41 ) /* Nao fazer ajuste contabil para LI */
		{
			ValorCheques += Doc.Valor;
		}
		else if( Doc.Vinculo == 0 && (Doc.TipoGenerico == "CO" || Doc.TipoDocto == 34) ) 
		{
			ValorContas += Doc.Valor;
		}
	}
	
	Diferenca = ValorCheques - ValorContas;

	/******************************************
	' * Oitava Fase do Vinculo                *
	' * Vinculando n Contas com n Cheques     *
	' * com diferenca <= ValorAjusteContabil  *
	' *****************************************/
	if( abs(Diferenca) >= 0 && 
		abs(Diferenca) <= Converte(m_pParametro->m_ValorAjusteContabil) )
	{
		iVinculo = 0;

		for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
		{
			Doc = m_ArrayDoc[i];
			if( (Doc.TipoGenerico == "CP" || Doc.TipoGenerico == "CO") 
				 && Doc.Vinculo == 0 && Doc.TipoDocto != 41)
			{
				/********************************
				' * Vinculando Contas e Cheques *
				' *******************************/
				if( iVinculo == 0 )
					iVinculo = Doc.IdDocto;

				m_ArrayDoc[i].Vinculo = iVinculo;
				/***********************************************
				 * Se cheque Ubb transforma em cheque terceiro *
				 ***********************************************/
				if( Doc.TipoDocto == 5 || Doc.TipoDocto == 7)
				{
					m_ArrayDoc[i].TipoDocto = 6;
					if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
					{
						m_ArrayDoc[i].Alcada = TRUE;
					}
					else
					{
						m_ArrayDoc[i].Alcada = FALSE;
					}
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}

		if( abs(Diferenca) > 0 )
		{
			/*****************************************
			' * Gravando Ajuste de Débito ou Crédito *
			' ****************************************/
			if( !InsereAjuste( (Diferenca > 0 ? 42 : 43), 
							   0, "0", abs(Diferenca), IdDocto ) )
			{
				return 0;
			}
			Ajuste.Alcada = FALSE;
			Ajuste.DesprezarVinculo = FALSE;
			Ajuste.IdDocto = IdDocto;
			Ajuste.TipoDocto = (Diferenca > 0 ? 42 : 43);
			Ajuste.TipoGenerico = (Diferenca > 0 ? "AC" : "AD");
			Ajuste.Valor = abs(Diferenca);
			Ajuste.Vinculo = iVinculo;
			m_ArrayDoc.InsertAt(0, Ajuste, 1);
		}
	}

	/*************************************
	  Verifica os titulos e cps do malote
	*************************************/
	lVinculo = VerificaVinculoMalote();
	if( lVinculo )
	{
		RemoveVinculo( lVinculo );
		RemoveAjustes();
	}
	// Sucesso !

	return 1;
}


void CVinculador::VinculaDocumentoEnvelope( void )
{
	DOCUMENTO Doc, DocAux, Ajuste;
	int i, j, iQtdCheques, iQtdContas;
	__int64 ValorVinculo;
	long IdDocto;
	int iConta;
	int iInicio, iDesprezar;
	CArray<int, int> aIndConta;
	
	/* Comentario do Vagner:
	   Como o vinculo do envelope eh um subconjunto do vinculo
	   do malote, este algoritmo foi feito baseado no algoritimo
	   do vinculo do malote. Por absoluta falta de tempo nao
	   foi possivel criar um melhor.
	*/

	/*****************************************
    ' * Marcar para Desprezar para o Vínculo *
    ' * Os Cheques com o mesmo Valor         *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		iQtdCheques = 0;
		iQtdContas  = 0;

		Doc = m_ArrayDoc[i];
		
		if( Doc.DesprezarVinculo == FALSE && Doc.Vinculo == 0 && 
			Doc.TipoDocto > 3 && Doc.TipoDocto <= 6 )
		{
            /****************************
            ' * Cheque/ADCC sem Vínculo *
            ' ***************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.Vinculo == 0 && DocAux.Valor == Doc.Valor )
				{
					if( DocAux.TipoDocto > 3 && DocAux.TipoDocto <= 6 )
					{
                        /*******************************
                        ' * Cheque/ADCC no Mesmo Valor *
                        ' ******************************/
						iQtdCheques++;
					}
					else if( DocAux.TipoGenerico == "CO" )
					{
                        /*************************
                        ' * Conta no mesmo Valor *
                        ' ************************/
						iQtdContas++;
					}
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}

		if( iQtdContas > 1 && m_QtdChequePagto > 1 )
		{
            /*************************************************
            ' * Desprezar para o Vínculo Automático          *
            ' ************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdContas > 1 && iQtdCheques > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdCheques > 1 && m_QtdContas > 1 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdCheques > 1 && iQtdContas > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}

		if( m_ArrayDoc[i].DesprezarVinculo )
		{
            /*****************************************
            ' * Marcar para Desprezar para o Vínculo *
            ' * Todos que tenham o mesmo valor       *
            ' ****************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				if( Doc.Valor == m_ArrayDoc[j].Valor )
					m_ArrayDoc[j].DesprezarVinculo = TRUE;
			}
		}
	}

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		iQtdContas     = 0;
		iQtdCheques    = 0;

		Doc = m_ArrayDoc[i];
		if( Doc.DesprezarVinculo == FALSE && Doc.Vinculo == 0 && 
			Doc.TipoGenerico == "CO" )
		{
            /***********************
            ' * Cheque sem Vínculo *
            ' **********************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.Vinculo == 0 && DocAux.Valor == Doc.Valor )
				{
					if( DocAux.TipoDocto > 3 && DocAux.TipoDocto <= 6 )
						iQtdCheques++;
					else if( DocAux.TipoGenerico == "CO" )
						iQtdContas++;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}


		if( iQtdCheques > 1 && m_QtdContas > 1 )
		{
            /*************************************************
            ' * Desprezar para o Vínculo Automático          *
            ' ************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdCheques > 1 && iQtdContas > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdContas > 1 && m_QtdChequePagto > 1 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}
		else if( iQtdContas > 1 && iQtdCheques > 0 )
		{
            /***********************************************************
            ' * Desprezar para o Vínculo Automático                    *
            ' **********************************************************/
			m_ArrayDoc[i].DesprezarVinculo = TRUE;
		}

		if( m_ArrayDoc[i].DesprezarVinculo )
		{
            /*****************************************
            ' * Marcar para Desprezar para o Vínculo *
            ' * Todos que tenham o mesmo valor       *
            ' ****************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				if( Doc.Valor == m_ArrayDoc[j].Valor )
					m_ArrayDoc[j].DesprezarVinculo = TRUE;
			}
		}
		
	}

    /****************************************
    ' * Primeira Fase do Vinculo            *
    ' * Vinculando Um Cheque para Uma Conta *
    ' ***************************************/

	/* Comentario do Vagner:
	   Nesta fase, se o Cheque de Pagamento for do Unibanco
	   pode-se vincula-lo a qualquer Cobranca/Arrecadacao.
	   Se o Cheque de Pagamento nao for do Unibanco
	   somente vincula-o se a Cobranca for diferente
	   de "Titulos Terceiros Sem CB" (12) e diferente de
	   "Cobranca Terceiros" (31)
	*/

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CP" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /*****************************************
            ' * Cheque ou ADCC a Verificar o Vinculo *
            ' ****************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.TipoGenerico == "CO" && 
					DocAux.Vinculo == 0 && 
					!DocAux.DesprezarVinculo &&
					Doc.Valor == DocAux.Valor )
				{
					/************************************
					  Versao 2.7
					  Considerar Titulos do Bandeirantes 
					  olhando para o codigo de barras
					************************************/
					if( (Doc.TipoDocto == 6 && DocAux.TipoDocto != 12 && DocAux.TipoDocto != 31) || 
						 Doc.TipoDocto == 5 ||
						 Doc.TipoDocto == 4 || 
						 Doc.TipoDocto == 41 ||
						(DocAux.TipoDocto == 31 && DocAux.Leitura.Left(3) == _T("230")))
					{
						/*
						if( m_ArrayDoc[i].TipoDocto != 6 && (m_ArrayDoc[j].TipoDocto != 20 ||
															 m_ArrayDoc[j].TipoDocto != 21 ||
															 m_ArrayDoc[j].TipoDocto != 22 ||
															 m_ArrayDoc[j].TipoDocto != 23) )
						
						if( m_ArrayDoc[i].TipoDocto == 6 && (m_ArrayDoc[j].TipoDocto != 12 ||
															 m_ArrayDoc[j].TipoDocto != 31) )
						{
						*/
							/**********************************
							' * Vinculando Conta com o Cheque *
							' *********************************/
							m_ArrayDoc[i].Vinculo  = Doc.IdDocto;
							m_ArrayDoc[j].Vinculo  = Doc.IdDocto;
							/**************************************
							' * Corrige o tipodocumento do cheque *
							' *************************************/
							if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
							{
								if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230"))
								{
									m_ArrayDoc[i].TipoDocto = 5;
									if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
									{
										m_ArrayDoc[i].Alcada = TRUE;
									}
									else
									{
										m_ArrayDoc[i].Alcada = FALSE;
									}
								}
								else
								{
									m_ArrayDoc[i].TipoDocto = 6;
									if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
									{
										m_ArrayDoc[i].Alcada = TRUE;
									}
									else
									{
										m_ArrayDoc[i].Alcada = FALSE;
									}
								}
							}
							break;
						/*}*/
					}
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

    /*********************************************
    ' * Verificando a Qtde de Contas no Envelope *
    ' * A Serem Vinculadas                       *
    ' *******************************************/
    iQtdContas = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CO" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdContas++;
	}

	if( iQtdContas == 0 )
	{
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return;
	}

	/* Comentario do Vagner:
	   Nesta segunda fase o algoritmo tenta vincular um cheque com
	   uma combinacao das contas, esta combinacao comeca com n, depois
	   tenta n-1, n-2, ...
	   Porem ele so testa combinacoes com documentos adjacentes, 
	   ele nao verifica todas as combinacoes possiveis.
	*/

	/* Comentario do Vagner:
	   Nesta fase, se, e somente se, o Cheque de Pagamento for 
	   do Unibanco pode-se vincula-lo a qualquer Cobranca/Arrecadacao.
	*/

    /*****************************************
    ' * Segunda Fase do Vinculo              *
    ' * Vinculando Um Cheque a várias Contas *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto == 4 || Doc.TipoDocto == 5 || Doc.TipoDocto == 41) && 
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /*****************************************
            ' * Cheque ou ADCC a Verificar o Vinculo *
            ' ****************************************/
			iInicio    = 1;
			iDesprezar = 1;
			while( iInicio <= iQtdContas )
			{
				ValorVinculo = 0;
				iConta       = 0;

				aIndConta.RemoveAll();
				for( j = 0; j < iQtdContas; j++ )
					aIndConta.Add(-1);

				for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux = m_ArrayDoc[j];
					if( DocAux.TipoGenerico == "CO" && 
						DocAux.Vinculo == 0 && 
						!DocAux.DesprezarVinculo )
					{
						iConta++;
						if( iConta == iInicio || iConta > iDesprezar )
						{
							ValorVinculo += DocAux.Valor;
							aIndConta[iConta-1] = j;
							if( ValorVinculo >= Doc.Valor )
								break;
						}
					}
				}
				if( iInicio % 50 == 0 )
				{
					PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
				}

				/**********************************************
				 * Verifica se o Valor da Combinacao bate com *
				 * o valor do cheque, entao vincula           *
				 **********************************************/
				if( ValorVinculo == Doc.Valor )
				{
					m_ArrayDoc[i].Vinculo = Doc.IdDocto;
					/**************************************
					' * Corrige o tipodocumento do cheque *
					' *************************************/
					if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
					{
						if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230") )
						{
							m_ArrayDoc[i].TipoDocto = 5;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
						else
						{
							m_ArrayDoc[i].TipoDocto = 6;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
					}

					for( j = 0; j < iQtdContas; j++ )
					{
						if( aIndConta[j] >= 0 )
						{
							m_ArrayDoc[aIndConta[j]].Vinculo = Doc.IdDocto;
						}
					}
					break;
				}

				iDesprezar++;
				if( iDesprezar >= iQtdContas )
				{
					iInicio++;
					iDesprezar = iInicio;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

	/* Comentario do Vagner:
	   Se o processo chegou ate este ponto, entao tudo que poderia
	   ser vinculado 1 para 1 ou 1 para N ja foi vinculado.
	*/

    /**********************************************
    ' * Verificando a Qtde de Cheques no Envelope *
    ' * A Serem Vinculados                        *
    ' ********************************************/
    iQtdCheques = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto >= 4 && Doc.TipoDocto < 7 || Doc.TipoDocto == 41) &&
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdCheques++;
	}

	if( iQtdCheques == 0 )
	{
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return;
	}

	/* Comentario do Vagner:
	   Nesta parte, o processo vai tentar vincular um cheque para uma
	   conta mesmo que os valores nao batam, desde que a diferenca seja
	   menor ou igual a ValorAjusteContabil, e caso seja possivel vincular,
	   sera gerado o ajuste contabil.
	*/

    /*****************************************
    ' * Terceira Fase do Vinculo             *
    ' * Vinculando Um Cheque para Uma Conta  *
	' * com diferenca <= ValorAjusteContabil *
    ' ****************************************/

	/* Comentario do Vagner:
	   Nesta fase, se o Cheque de Pagamento for do Unibanco
	   pode-se vincula-lo a qualquer Cobranca/Arrecadacao.
	   Se o Cheque de Pagamento nao for do Unibanco
	   somente vincula-o se a Cobranca for diferente
	   de "Titulos Terceiros Sem CB" (12) e diferente de
	   "Cobranca Terceiros" (31)
	*/

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CP" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /*****************************************
            ' * Cheque ou ADCC a Verificar o Vinculo *
            ' ****************************************/
			for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
			{
				DocAux = m_ArrayDoc[j];
				if( DocAux.TipoGenerico == "CO" && 
					DocAux.Vinculo == 0 && 
					!DocAux.DesprezarVinculo &&
					abs(Doc.Valor - DocAux.Valor) > 0 &&
				    abs(Doc.Valor - DocAux.Valor) <= Converte(m_pParametro->m_ValorAjusteContabil) )
				{
					/************************************
					  Versao 2.7
					  Considerar Titulos do Bandeirantes 
					  olhando para o codigo de barras
					************************************/
					if( (Doc.TipoDocto == 6 && DocAux.TipoDocto != 12 && DocAux.TipoDocto != 31) || 
						 Doc.TipoDocto == 5 ||
						 Doc.TipoDocto == 4 || 
						 Doc.TipoDocto == 41 ||
						(DocAux.TipoDocto == 31 && DocAux.Leitura.Left(3) == _T("230")))
					{
						/**********************************
						' * Vinculando Conta com o Cheque *
						' *********************************/
						m_ArrayDoc[i].Vinculo  = Doc.IdDocto;
						m_ArrayDoc[j].Vinculo  = Doc.IdDocto;
						/**************************************
						' * Corrige o tipodocumento do cheque *
						' *************************************/
						if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
						{
							if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230") )
							{
								m_ArrayDoc[i].TipoDocto = 5;
								if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
								{
									m_ArrayDoc[i].Alcada = TRUE;
								}
								else
								{
									m_ArrayDoc[i].Alcada = FALSE;
								}
							}
							else
							{
								m_ArrayDoc[i].TipoDocto = 6;
								if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
								{
									m_ArrayDoc[i].Alcada = TRUE;
								}
								else
								{
									m_ArrayDoc[i].Alcada = FALSE;
								}
							}
						}

						/*****************************************
						' * Gravando Ajuste de Débito ou Crédito *
						' ****************************************/
						if( !InsereAjuste( (Doc.Valor > DocAux.Valor ? 42 : 43), 
										   0, "0", abs(Doc.Valor - DocAux.Valor), IdDocto ) )
						{
							return;
						}
						Ajuste.Alcada = FALSE;
						Ajuste.DesprezarVinculo = FALSE;
						Ajuste.IdDocto = IdDocto;
						Ajuste.TipoDocto = (Doc.Valor > DocAux.Valor ? 42 : 43);
						Ajuste.TipoGenerico = (Doc.Valor > DocAux.Valor ? "AC" : "AD");
						Ajuste.Valor = abs(Doc.Valor - DocAux.Valor);
						Ajuste.Vinculo = Doc.IdDocto;
						m_ArrayDoc.InsertAt(0, Ajuste, 1);

						break;
					}
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

    /*********************************************
    ' * Verificando a Qtde de Contas no Malote   *
    ' * A Serem Vinculadas                       *
    ' *******************************************/
    iQtdContas = 0;

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( Doc.TipoGenerico == "CO" && Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
			iQtdContas++;
	}

	if( iQtdContas == 0 )
	{
        /******************************************
        ' * Não Existe Mais Documentos a Vincular *
        ' *****************************************/
		return;
	}

	/* Comentario do Vagner:
	   Nesta fase, se, e somente se, o Cheque de Pagamento for 
	   do Unibanco pode-se vincula-lo a qualquer Cobranca/Arrecadacao.
	*/

    /*****************************************
    ' * Quarta Fase do Vinculo               *
    ' * Vinculando Um Cheque a várias Contas *
	' * com diferenca <= ValorAjusteContabil *
    ' ****************************************/
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		if( (Doc.TipoDocto == 4 || Doc.TipoDocto == 5 || Doc.TipoDocto == 41) && 
			Doc.Vinculo == 0 && !Doc.DesprezarVinculo )
		{
            /*********************************************
            ' * Cheque, LI ou ADCC a Verificar o Vinculo *
            ' ********************************************/
			iInicio    = 1;
			iDesprezar = 1;
			while( iInicio <= iQtdContas )
			{
				ValorVinculo = 0;
				iConta       = 0;

				aIndConta.RemoveAll();
				for( j = 0; j < iQtdContas; j++ )
					aIndConta.Add(-1);

				for( j = 0; j <= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux = m_ArrayDoc[j];
					if( DocAux.TipoGenerico == "CO" && 
						DocAux.Vinculo == 0 && 
						!DocAux.DesprezarVinculo )
					{
						iConta++;
						if( iConta == iInicio || iConta > iDesprezar )
						{
							ValorVinculo += DocAux.Valor;
							aIndConta[iConta-1] = j;
							if( abs(Doc.Valor - ValorVinculo) <= 
								Converte(m_pParametro->m_ValorAjusteContabil) )
								break;
						}
					}
				}
				if( iInicio % 50 == 0 )
				{
					PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
				}

				/*******************************************************
				 * Verifica se o Valor da Combinacao menos             *
				 * o valor do cheque, e menor que ValorAjusteContabil  *
				 *******************************************************/
				if( abs(Doc.Valor - ValorVinculo) > 0 &&
					abs(Doc.Valor - ValorVinculo) <= 
					Converte(m_pParametro->m_ValorAjusteContabil) )
				{
					m_ArrayDoc[i].Vinculo = Doc.IdDocto;
					/**************************************
					' * Corrige o tipodocumento do cheque *
					' *************************************/
					if( m_ArrayDoc[i].TipoDocto == 6 || m_ArrayDoc[i].TipoDocto == 7 )
					{
						if( m_ArrayDoc[i].Leitura.Left(3) == _T("409") || m_ArrayDoc[i].Leitura.Left(3) == _T("230") )
						{
							m_ArrayDoc[i].TipoDocto = 5;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcada_Mal : m_pParametro->m_ValorAlcada_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
						else
						{
							m_ArrayDoc[i].TipoDocto = 6;
							if(	m_ArrayDoc[i].Valor >= Converte((m_Capa.IdEnv_Mal == "M" ? m_pParametro->m_ValorAlcadaOutros_Mal : m_pParametro->m_ValorAlcadaOutros_Env)) )
							{
								m_ArrayDoc[i].Alcada = TRUE;
							}
							else
							{
								m_ArrayDoc[i].Alcada = FALSE;
							}
						}
					}

					for( j = 0; j < iQtdContas; j++ )
					{
						if( aIndConta[j] >= 0 )
						{
							m_ArrayDoc[aIndConta[j]].Vinculo = Doc.IdDocto;
						}
					}

					/*****************************************
					' * Gravando Ajuste de Débito ou Crédito *
					' ****************************************/
					if( !InsereAjuste( (Doc.Valor > ValorVinculo ? 42 : 43), 
									   0, "0", abs(Doc.Valor - ValorVinculo), IdDocto ) )
					{
						return;
					}
					Ajuste.Alcada = FALSE;
					Ajuste.DesprezarVinculo = FALSE;
					Ajuste.IdDocto = IdDocto;
					Ajuste.TipoDocto = (Doc.Valor > ValorVinculo ? 42 : 43);
					Ajuste.TipoGenerico = (Doc.Valor > ValorVinculo ? "AC" : "AD");
					Ajuste.Valor = abs(Doc.Valor - ValorVinculo);
					Ajuste.Vinculo = Doc.IdDocto;
					m_ArrayDoc.InsertAt(0, Ajuste, 1);
					
					break;
				}

				iDesprezar++;
				if( iDesprezar >= iQtdContas )
				{
					iInicio++;
					iDesprezar = iInicio;
				}
			}
			if( i % 10 == 0 )
			{
				PeekMessage(&m_Msg, NULL, 0, 0, PM_NOREMOVE);
			}
		}
	}

}


BOOL CVinculador::AtualizaDocumentos( void )
{
	CString strSql;
	int i;
	DOCUMENTO Doc;

	m_oDB.BeginTrans();

	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++ )
	{
		Doc = m_ArrayDoc[i];
		
		strSql.Format("Execute VA_AtualizaDocumento %ld, %ld, %d, %d, '%s'",
					  m_lDataProc, Doc.IdDocto, Doc.TipoDocto, 
					  Doc.Vinculo, (Doc.Alcada ? "S" : "N") );

		try
		{
			m_oDB.ExecuteSQL(LPCTSTR(strSql));
		}
		catch (CDBException *E)
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, LPCSTR(E->m_strError)); 
			strcat(m_MsgError, " (CGetDataProc.Open) ");
			m_iCodError= E->m_nRetCode;  

			m_oDB.Rollback();

			E->Delete(); 
			return FALSE;		
		}

	}

	m_oDB.CommitTrans();

	return TRUE;

}

BOOL CVinculador::RemoveAjustes( void )
{
	CString strSql;

	strSql.Format("Execute RemoveAjusteCapa %ld, %ld", m_lDataProc, m_Capa.IdCapa);

	try
	{
		m_oDB.ExecuteSQL(LPCTSTR(strSql));
	}
	catch (CDBException *E)
	{
		//Armazena a informação sobre o erro
		strcpy(m_MsgError, LPCSTR(E->m_strError)); 
		strcat(m_MsgError, " (RemoveAjusteCapa) ");
		m_iCodError= E->m_nRetCode;  

		E->Delete(); 
		return FALSE;		
	}
	return TRUE;
}

BOOL CVinculador::DespachaCapa( void )
{
	BOOL bAlcada, bVinculado, bDifDeposito, bNovaRegraLI;
	__int64 ValorCredito, ValorDebito, ValorChDeposito, ValorDeposito;
	DOCUMENTO Doc, DocAux, DocAux2;
	int i, j, k;

	ValorCredito	= 0;
	ValorDebito		= 0;
	ValorChDeposito = 0;
	ValorDeposito   = 0;
	bAlcada			= FALSE;
	bVinculado		= TRUE;
	bDifDeposito	= FALSE;
	bNovaRegraLI	= FALSE;

	i = 0;
	while( i <= m_ArrayDoc.GetUpperBound() )
	{
		Doc = m_ArrayDoc[i];
		
		if( Doc.TipoGenerico == "CO" || Doc.TipoDocto == 34 || Doc.TipoDocto == 42 )
		{
			ValorCredito += Doc.Valor;

		}
		else if( Doc.TipoGenerico == "CP" || Doc.TipoDocto == 38 || Doc.TipoDocto == 43 )
		{
			ValorDebito += Doc.Valor;
		}
		else if( Doc.TipoGenerico == "DE" )
		{

			ValorDeposito += Doc.Valor;

			/**************************************
			  Verifica se está na Nova Regra do LI
			**************************************/
			if( bNovaRegraLI == FALSE )
			{
				j = i + 1;

				DocAux = m_ArrayDoc[j];

				while( j <= m_ArrayDoc.GetUpperBound() )
				{
					if( m_ArrayDoc[j].TipoGenerico == "DE" || m_ArrayDoc[j].TipoGenerico == "OC" )
					{
						ValorDeposito += m_ArrayDoc[j].Valor;
						j++;
						continue;
					}
					/******************************
					  É um LI seguido de Depositos
					******************************/
					else if( m_ArrayDoc[j].TipoGenerico == "CP" && m_ArrayDoc[j].TipoDocto == 41 )
					{

						/*************************************************************
						 conta e soma os Li's na Capa - alteração aceitar varios LI's
						**************************************************************/
						while(j <= m_ArrayDoc.GetUpperBound() && m_ArrayDoc[j].TipoDocto == 41 )
						{
							ValorChDeposito += m_ArrayDoc[j].Valor;
							if( m_ArrayDoc[j].Alcada )
								bAlcada = TRUE;

							if( m_ArrayDoc[j].Vinculo == 0 )
								bVinculado = FALSE;
							j ++ ;
						}
						bNovaRegraLI = TRUE;

						k = j;
						/******************************
						  Procura agora somente contas
						 ******************************/

						while( k <= m_ArrayDoc.GetUpperBound() )
						{
							if( (m_ArrayDoc[k].TipoGenerico != "CO") &&
								(m_ArrayDoc[k].TipoGenerico != "AD") &&
								(m_ArrayDoc[k].TipoGenerico != "AC"))
							{
								bNovaRegraLI = FALSE;
								break;
							}
							else
							{
								if( m_ArrayDoc[k].TipoGenerico == "AD" )
									ValorCredito -= m_ArrayDoc[k].Valor;
								else
									ValorCredito += m_ArrayDoc[k].Valor;
							}
							k++;
						}
					}
					/*************************
					  Não está na Regra do LI
					*************************/
					break;
				}
			}

			if( bNovaRegraLI == FALSE )
			{
				for( j = i + 1; j<= m_ArrayDoc.GetUpperBound(); j++ )
				{
					DocAux2 = m_ArrayDoc[j];

					if( DocAux2.TipoGenerico == "CD" ||
						DocAux2.TipoDocto == 33 || DocAux2.TipoDocto == 41 ) 
					{
						ValorChDeposito += DocAux2.Valor;
					}
					else if( DocAux2.TipoDocto == 32 || ( m_Capa.CSP == TRUE && DocAux2.TipoGenerico == "DE" ) )
					{
						ValorDeposito += DocAux2.Valor;
					}
					else if( m_Capa.CSP == TRUE )
					{
						if( DocAux2.TipoGenerico == "AC" )
						{
							ValorDeposito += DocAux2.Valor;
						}
						else if( DocAux2.TipoGenerico == "AD" )
						{
							ValorChDeposito += DocAux2.Valor;
						}
					}
					else
					{
						break;
					}
				}
				if( !bDifDeposito && ValorChDeposito != ValorDeposito )
				{
					bDifDeposito = TRUE;
				}
			}
			else
			{
				if( !bDifDeposito && (ValorChDeposito + ValorDebito) != (ValorDeposito + ValorCredito) )
				{
					bDifDeposito = TRUE;
				}
			}
			i = j - 1;
		}
		
		if( Doc.Alcada )
			bAlcada = TRUE;

		if( Doc.Vinculo == 0 )
			bVinculado = FALSE;

		i++;

		if( bNovaRegraLI ) break;
	}

	m_Capa.Diferenca = (ValorChDeposito + ValorDebito) - (ValorDeposito + ValorCredito);

	if( m_Capa.Diferenca != 0 )
	{
		if( m_Capa.CSP == TRUE )
		{
			//Enviar para CSP Novamente
			m_Capa.Status = "N";
		}
		else if( !m_Capa.GerarAjuste ) //|| bNovaRegraLI)
		{
			// Enviar para Prova Zero
			m_Capa.Status = "4";
		}
		else
		{
			// Enviar para Vinculo Manual
			m_Capa.Status = "7";
		}
	}
	else if( m_Capa.Diferenca == 0 && bDifDeposito )
	{
		if( m_Capa.CSP == TRUE )
		{
			//Enviar para CSP Novamente
			m_Capa.Status = "N";
		}
		else
		{
			// Enviar para Vinculo Automatico
			m_Capa.Status = "9";
		}
	}
	else if( !bVinculado )
	{
		// Enviar para Vinculo Manual
		m_Capa.Status = "7";
	}
	else if( bAlcada )
	{
		if( m_Capa.CSP == TRUE )
		{
			m_Capa.Status = "R";
		}
		else
		{
			// Enviar para Alcada
			m_Capa.Status = "6";
		}
	}
	else
	{
		// Enviar para Transmissao
		m_Capa.Status = "R";
	}

	if( AtualizaStatusCapa() )
		return TRUE;
	else
		return FALSE;
}

__int64 CVinculador::Converte( CString Valor )
{
	CString Result;
	int Pos;

	Pos = Valor.Find( '.' );
	if( Pos == -1 )
	{
		Result = Valor;
	}
	else
	{
		Result = Valor.Left(Pos) + Valor.Mid(Pos + 1, 2);
	}
	return _atoi64(LPCTSTR(Result));
}

BOOL CVinculador::InsereLog( short Acao )
{
	CString strSql;

	strSql.Format("Execute InsereLog %ld, %ld, 0, 'Vinculo', %d",
				  m_lDataProc, m_Capa.IdCapa, Acao);

	try
	{
		m_oDB.ExecuteSQL(LPCTSTR(strSql));
	}
	catch (CDBException *E)
	{
		//Armazena a informação sobre o erro
		strcpy(m_MsgError, LPCSTR(E->m_strError)); 
		strcat(m_MsgError, " (InsereLog) ");
		m_iCodError= E->m_nRetCode;  

		E->Delete(); 
		return FALSE;		
	}
	return TRUE;
}

BOOL CVinculador::PossuiDocumentoTransmitido( void )
{
	CGetDocumentoTransmitido m_DocTransmitido(&m_oDB);

	while( true )
	{
		try
		{
			if( m_DocTransmitido.IsOpen() )
				m_DocTransmitido.Close();

			m_DocTransmitido.m_DataProc = m_lDataProc;
			m_DocTransmitido.m_IdCapa   = m_Capa.IdCapa;
			if( !m_DocTransmitido.Open( CRecordset::snapshot, NULL ) )
			{
				//Armazena a informação sobre o erro
				strcpy(m_MsgError, "Erro na obtenção na Qtde de Documentos Transmitidos. (CGetDocumentoTransmitido.Open)"); 
				m_iCodError= 100;  
				return FALSE;
			}
			if( m_DocTransmitido.m_Qtde > 0)
			{

				/*
					se não veio de CSP
				*/
				if(! CapaCSP() )
				{
					m_Capa.Status = "V";
					m_DocTransmitido.Close();
					return TRUE;
				}
				else
				{
					m_DocTransmitido.Close();
					return FALSE;
				}
				/*=======================================
				  O número de documentos transmitidos é
				  igual ao número de documentos na capa

				 *Enviar para verificação
				=======================================*/
/*
				if( m_DocTransmitido.m_QtdeDoctosDaCapa == m_DocTransmitido.m_Qtde )
				{
					// So exitem documentos transmitidos
					// Enviar para verificacao
					m_Capa.Status = "V";
					m_DocTransmitido.Close();
					return TRUE;
				}
				else
				{
					m_DocTransmitido.Close();
					return FALSE;
				}
*/
			}
			else
			{
				m_DocTransmitido.Close();
				return FALSE;
			}
		}
		catch(CDBException *E)
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, LPCSTR(E->m_strError)); 
			strcat(m_MsgError, " (CGetDocumentoTransmitido.Open) ");
			m_iCodError= E->m_nRetCode;  
			E->Delete(); 

			if( memcmp(m_MsgError, "Timeout expired", 15) != 0 )
			{
				return FALSE;		
			}
			else
				continue;
		}
	}
}


int CVinculador::NovaRegraLancamentoInterno( void )
{

/*********************************************************************************
	RETORNO:
		TRUE  - Não contem documentos a não ser Nova regra do Lancamento Interno
		FALSE - Não está na nova regra do Lancamento Interno
		> 0   - O indice do primeiro deposito da regra do Lancamento Interno
*********************************************************************************/

	int       i, retorno;
	i             = 0;

	/******************
	 Default de retorno
     ******************/
	retorno = FALSE;

	while( i <= m_ArrayDoc.GetUpperBound() )
	{
		if( m_ArrayDoc[i].TipoGenerico == "DE" || m_ArrayDoc[i].TipoGenerico == "OC" )
		{
			i++;
			continue;
		}
		else if ( m_ArrayDoc[i].TipoGenerico == "CP" && m_ArrayDoc[i].TipoDocto == 41 )
		{

			while(i <= m_ArrayDoc.GetUpperBound() && m_ArrayDoc[i].TipoDocto == 41 )
			{
				retorno = -1;
				i ++ ;
			}

			while( i <= m_ArrayDoc.GetUpperBound() )
			{
				if( (m_ArrayDoc[i].TipoGenerico != "CO") &&
					(m_ArrayDoc[i].TipoGenerico != "AD") &&
					(m_ArrayDoc[i].TipoGenerico != "AC"))
				{
					retorno = FALSE;
					break;
				}
				i++;
			}
			
		}
		else 
		{
			break;
		}

	i ++;

	}

	return retorno;
}



BOOL CVinculador::CapaCSP ( void )
{
	CGetControleCapa m_ControleCapa(&m_oDB);

	m_Capa.CSP = FALSE;

	while( true )
	{
		try
		{
			if( m_ControleCapa.IsOpen() )
				m_ControleCapa.Close();

			m_ControleCapa.m_DataProc = m_lDataProc;
			m_ControleCapa.m_IdCapa   = m_Capa.IdCapa;
			if( !m_ControleCapa.Open( CRecordset::snapshot, NULL ) )
			{
				/*===================================
				  Se não conseguiu abrir,
				  armazena informações sobre o erro
				===================================*/
				strcpy(m_MsgError, "Erro na obtenção do Controle Capa. (CGetControleCapa.Open)");
				m_iCodError = 100;
				return FALSE;
			}
			if( m_ControleCapa.m_IdModulo > 0 )
			{
				//Verifica se é de CSP
				if( m_ControleCapa.m_IdModulo == MODULO_CSP )
				{
					m_Capa.CSP = TRUE;
					m_ControleCapa.Close();
					return TRUE;
				}
				else
				{
					m_ControleCapa.Close();
					return FALSE;
				}
			}
			else
			{
				m_ControleCapa.Close();
				return FALSE;
			}
		}
		catch(CDBException *E)
		{
			//Armazena a informação sobre o erro
			strcpy(m_MsgError, LPCSTR(E->m_strError));
			strcat(m_MsgError, " (CGetControleCapa.Open) ");
			m_iCodError = E->m_nRetCode;
			E->Delete();

			if( memcmp(m_MsgError, "Timeout expired", 15) != 0 )
			{
				return FALSE;
			}
			else
				continue;
		}
	}

	return FALSE;
}

/*===================================================
  ContemAjusteVinculo:
  Retorna TRUE se o determinado vinculo tem ajuste, 
  caso contrário retorno é FALSE
===================================================*/
BOOL CVinculador::ContemAjusteVinculo( long pVinculo )
{
	int i = 0;
	BOOL bRetorno = FALSE;

	while( i <= m_ArrayDoc.GetUpperBound() )
	{
		if( m_ArrayDoc[i].Vinculo == pVinculo )
		{
			if( m_ArrayDoc[i].TipoGenerico == "AD" ||
				m_ArrayDoc[i].TipoGenerico == "AC" )
			{
				bRetorno = TRUE;
				break;
			}
		}
		i++;
	}

	return bRetorno;

}

void CVinculador::RemoveVinculo(long pVinculo)
{
	int i;
	int index;

	//Remove vinculo
	index = -1;
	for( i = 0; i <= m_ArrayDoc.GetUpperBound(); i++)
	{
		if( m_ArrayDoc[i].Vinculo == pVinculo )
		{
			if( m_ArrayDoc[i].TipoGenerico == "CP" ||
				m_ArrayDoc[i].TipoGenerico == "CO" )
			{
				m_ArrayDoc[i].Vinculo = 0;
			}
			else if( m_ArrayDoc[i].TipoGenerico == "AC" ||
					 m_ArrayDoc[i].TipoGenerico == "AD" )
			{
				m_ArrayDoc[i].Vinculo = 0;
				index = i;
			}
		}
	}

	//por ultimo remove o AD ou AC
	if( index != -1 )
	{
		m_ArrayDoc.RemoveAt(index,1);
	}
}

long CVinculador::VerificaVinculoMalote( void )
{
	long lUltVinculoProc, lVinculo;
	int i, j;
	BOOL ContemCPTerceiro = FALSE;
	BOOL ContemCOTerceiro = FALSE;
	DOCUMENTO Doc;

	/*******************************************
	  Verificar se são titulos de outros bancos
	  e contas de outros bancos
	*******************************************/

	lUltVinculoProc = 0;
	if( m_Capa.CSP == FALSE )
	{
		j = 0;
		while( j <= m_ArrayDoc.GetUpperBound() )
		{
			Doc = m_ArrayDoc[j];

			lVinculo = Doc.Vinculo;
			if( lVinculo != lUltVinculoProc )
			{
				lUltVinculoProc = lVinculo;
				i = j;
				while( i <= m_ArrayDoc.GetUpperBound() )
				{
					Doc = m_ArrayDoc[i];
					if( lVinculo == Doc.Vinculo )
					{
						/*
						
						if( Doc.TipoGenerico == "CP" && Doc.TipoDocto == 6 )
						{
							ContemCPTerceiro = TRUE;
						}
						else if( Doc.TipoGenerico == "CO" && (Doc.TipoDocto != 28 &&
															  Doc.TipoDocto != 29 &&
															  Doc.TipoDocto != 30))
						{
							ContemCOTerceiro = TRUE;
						}
						*/

						//Retirado 'if' acima para permitir vinculo de LI com cheque 3o.
						if( Doc.TipoGenerico == "CO" && (Doc.TipoDocto != 28 &&
														 Doc.TipoDocto != 29 &&
														 Doc.TipoDocto != 30))
						{
							ContemCOTerceiro = TRUE;
						}
					}
					i++;
				}
			}
			/************************************
			  Se em algum vinculo da capa conter
			  Conta e Cheque de terceiro enviar
			  para Vinculo Manual
			************************************/
			if( ContemCPTerceiro && ContemCOTerceiro )
			{
				return lVinculo;
			}
			j++;
		}
	}
	
	return FALSE;

}