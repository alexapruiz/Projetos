// VipsDll.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "stdlib.h"
#include "stdio.h"
#include "direct.h"
#include "errno.h"
#include "VipsDll.h"
#include "VipsDrv.h"
#include "VipsGrim.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
			m_ComPort    = 1;             // Serial 1
			m_Boxes      = 1;             // Nro. de Escanhinhos
			m_Resolution = RES100DPI;     // 100 Dpi
			m_Reader     = 3;             // Ambas
			m_ImageType  = 3;             // JPG
			m_Compress   = 30;            // Fator de Compressao
			m_Threshold  = 50;            // Fator para corte das imagens
			m_Directory[0]  = '\0';       // Diretorio das imagens
			m_CameraFile[0] = '\0';       // Arq. conf. camera
			m_Sequencial = 1;             // Seq. da imagem
			m_Box        = 1;             // Escaninho atual
			m_DocBox     = 0;             // Qtde documentos no escaninho atual
			m_MaxDocBox  = MAX_DOC_BOX;   // Qtde maxima de docs por escaninho
			m_BoxDefault = 0;             // Escaninho default : 0 = Nenhum
			break;
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}

// Inicializa Dll's 
long WINAPI VIPS_Init( void )
{
	return Init();
}

// Libera as Dll's
void WINAPI VIPS_Done( void )
{
	Done();
}

// Seleciona a porta serial
void WINAPI VIPS_SetComPort( long ComPort )
{
	m_ComPort = ComPort;
}

// Seleciona a resolucao ( 100 ou 200 DPI )
void WINAPI VIPS_SetResolution( long Res )
{
	if( Res == 200 )
		m_Resolution = RES200DPI;
	else
		m_Resolution = RES100DPI;
}

// Seleciona leitora ( 1 = CMC7; 2 = Cod. Barras; 3 = Ambas )
void WINAPI VIPS_SetReader( long Reader )
{
	m_Reader = Reader;
}

// Seleciona a quantidade de Escaninhos ( Nro. de escaninhos deve ser impar )
void WINAPI VIPS_SetBoxes( long Boxes )
{
	short i, Box, Metade;

	if( Boxes <= MAX_BOXES )
		m_Boxes = Boxes;
	else
		m_Boxes = MAX_BOXES;

	Box = 1;
	Metade = (short)(m_Boxes / 2);
	for( i = 1; i <= Metade; i++ )
	{
		m_aBoxes[i]        = Box;
		m_aBoxes[i+Metade] = Box + 1;
		Box += 2;
	}
	m_aBoxes[m_Boxes] = (short)m_Boxes;
}

// Seleciona a quantidade maxima de docs por escanhinho
void WINAPI VIPS_SetMaxDocBox( long MaxDocBox )
{
	m_MaxDocBox = MaxDocBox;
}

// Seleciona o escanhinho default ( Default = 0 : Nenhum )
void WINAPI VIPS_SetBoxDefault( long BoxDefault )
{
	if( BoxDefault >= 0 && BoxDefault <= m_Boxes )
		m_BoxDefault = (short)BoxDefault;
}

// Seleciona o tipo de imagem gerada ( 1 = BMP; 2 = TIF; 3 = JPG )
void WINAPI VIPS_SetImageType( long ImageType )
{
	m_ImageType = ImageType;
}

// Seleciona o fator de compressao do JPG ( defualt = 30 ) 
void WINAPI VIPS_SetCompress( long Fator )
{
	m_Compress = Fator;
}

// Seleciona se deve cortar as bordas 
// ( Valor-> 0 a 255, Zero nao corta, default = 50. Quanto mais alto, mais corta )
void WINAPI VIPS_SetCutBords( long Valor )
{
	m_Threshold = Valor;
}

// Seleciona o diretorio onde serao gravadas as imagens
void WINAPI VIPS_SetImageDirectory( char *Diretorio )
{
	strcpy(m_Directory, Diretorio);
}

// Seleciona o arq de conf. da camera
void WINAPI VIPS_SetCameraFile( char *CameraFile )
{
	strcpy(m_CameraFile, CameraFile);
}

// Executa a captura
long WINAPI VIPS_Captura( long AgProc, long Lote, char *ArqRetorno, long Append )
{
	long lRet;
	int  iRet;
	char Directory[256];

	sprintf(Directory, "%s%09.9ld", m_Directory, Lote);
	if( _mkdir(Directory) == -1 && errno == ENOENT )
	{
		// nao foi possivel criar diretorio do lote
		return -1004;
	}
	
	if( m_Lote != Lote || m_AgProc != AgProc)
	{
		Append = FALSE;
	}
	
	m_AgProc = AgProc;
	m_Lote   = Lote;

	if( !Append )
	{
		m_fData = fopen(ArqRetorno, "wb");

		GravaLog("VIPS_Captura - Criar arquivo de retorno\n");

		m_Sequencial = 1;
		m_Box = 1;
		m_DocBox = 0;

		sprintf(m_strLog, "VIPS_Captura - m_Box = %d, m_DocBox = %ld\n", m_Box, m_DocBox);
		GravaLog(m_strLog);
	}
	else
	{
		m_fData = fopen(ArqRetorno, "a+b");
	}
	
	if( m_fData == NULL )
	{
		// nao foi possivel abrir arquivo de dados
		return -1003;
	}

	GravaLog("VIPS_Captura - Inicio ProcessoCaptura\n");
	
	lRet = ProcessoCaptura();

	if( m_IsToInit )
	{
		TratamentoLimpaBuffer();
	}

	// aguarda a Callback gravar a ultima imagem
	while( m_Busy > 0 )
	{
		Sleep(500);
	}
	
	if( m_fData )
	{
		fflush(m_fData);
		fclose(m_fData);
		m_fData = NULL;
	}

	vipsRecupereStatusLibelle(m_strMsg, &iRet);
	iRet = vipsEnd();

	sprintf(m_strLog, "VIPS_Captura - vipsEnd = %d\n", iRet);
	GravaLog(m_strLog);

	if( m_IsToInit )
	{
		VIPS_Reset();

		m_IsToInit = false;
	}

	GravaLog("VIPS_Captura - Termino ProcessoCaptura\n");

	sprintf(m_strLog, "VIPS_Captura - lRet = %ld\n", lRet);
	GravaLog(m_strLog);

	return lRet;
}

// Executa a recaptura
long WINAPI VIPS_Recaptura( long AgProc, long Lote, long SeqInicial, char *ArqRetorno, long Append )
{
	long lRet;
	int  iRet;
	char Directory[256];

	sprintf(Directory, "%s%09.9ld", m_Directory, Lote);
	
	if( Append && ( m_Lote != Lote || m_AgProc != AgProc ) )
	{
		// Se eh para continuar no mesmo arquivo nao pode mudar o Lote e/ou Agencia
		return -1005;
	}

	m_AgProc = AgProc;
	m_Lote   = Lote;
	m_Box = 1;
	m_DocBox = 0;
	
	if( !Append )
	{
		m_Sequencial = SeqInicial;

		m_fData = fopen(ArqRetorno, "wb");

		GravaLog("VIPS_Recaptura - Criar arquivo de retorno\n");
	}
	else
	{
		m_fData = fopen(ArqRetorno, "a+b");

		GravaLog("VIPS_Recaptura - Abrir arquivo de retorno\n");
	}

	sprintf(m_strLog, "VIPS_Recaptura - m_Box = %d, m_DocBox = %ld\n", m_Box, m_DocBox);
	GravaLog(m_strLog);
	
	if( m_fData == NULL )
	{
		// nao foi possivel abrir arquivo de dados
		return -1003;
	}

	GravaLog("VIPS_Recaptura - Inicio ProcessoCaptura\n");
	
	lRet = ProcessoCaptura();

	if( m_IsToInit )
	{
		TratamentoLimpaBuffer();
	}

	// aguarda a Callback gravar a ultima imagem
	while( m_Busy > 0 )
	{
		Sleep(500);
	}
	
	if( m_fData )
	{
		fflush(m_fData);
		fclose(m_fData);
		m_fData = NULL;
	}

	vipsRecupereStatusLibelle(m_strMsg, &iRet);
	iRet = vipsEnd();

	sprintf(m_strLog, "VIPS_Recaptura - vipsEnd = %d\n", iRet);
	GravaLog(m_strLog);

	if( m_IsToInit )
	{
		VIPS_Reset();

		m_IsToInit = false;
	}

	GravaLog("VIPS_Recaptura - Termino ProcessoCaptura\n");

	sprintf(m_strLog, "VIPS_Recaptura - lRet = %ld\n", lRet);
	GravaLog(m_strLog);

	return lRet;
}

// Reseta as Dll's
long WINAPI VIPS_Reset( void )
{
    int iRet;

    m_IsToInit = false;

	vipsClose();

    GravaLog("Reset - vipsClose\n");

	// on ouvre la communication avec le lecteur
    iRet = vipsOpen(&m_DCB);

	sprintf(m_strLog, "Reset - vipsOpen - iRet = %d\n", iRet);
	GravaLog(m_strLog);

    if (iRet != VER_NOERROR)  // si problème d'initialisation de la DLL vipsdrv
    {
        vipsRecupereStatusLibelle(m_strMsg, &iRet);
        TratamentoErro();
		return iRet;
    }

	return 0;

}

long Init( void )
{
    int iRet;

    if( strlen(m_Directory) == 0 || strlen(m_CameraFile) == 0 ) return -1000;

#ifdef _DEBUG
	m_fLog = fopen("log.txt", "wt");
#endif

	pFunctionFiler_STDCALL = FonctionUserDataBufferFichiers;

	memset(&m_LC13, '\0', sizeof(TVipsLC13));

	iRet = VipsGrimLC13_Init_Acquisition(&m_LC13,
											VIDEO_CARD,
											NB_BUFFERS_DMA,
											NB_BUFFERS_IMAGES,
											NB_BUFFERS_FILES,
											m_CameraFile,
											m_Directory,
											m_Resolution,
											FICHIER_BMP,
											(unsigned int)8,
											(unsigned int)30,
											(unsigned char)135,
											ROTATION_90,
											AUCUNE_OPERATION,
											ROTATION_270H, AUCUNE_OPERATION,                    // ROTATION_270, SYMETRIE_H,
											FALSE, 56, 212, 0, 0,
											pFunctionFiler_STDCALL,
											sizeof(TVipsUserDataBufferFichiers),
											CORRECTION_LUMINEUSE,
											CORRECTION_VERTICALE,
											ETALON_FILE);

	sprintf(m_strLog, "Init - VipsGrimLC13_Init_Acquisition - iRet = %d\n", iRet);
	GravaLog(m_strLog);

	sprintf(m_strLog, "Init - Diretorio Imagens = %s\n", m_Directory);
	GravaLog(m_strLog);
	sprintf(m_strLog, "Init - Camera File = %s\n", m_CameraFile);
	GravaLog(m_strLog);
	sprintf(m_strLog, "Init - Resolution = %c\n", m_Resolution);
	GravaLog(m_strLog);

    
	if (iRet)
	{
        TratamentoErro();
		return iRet;
	}

    // arquivo de dados
	m_fData = NULL;

	m_Lote = 0;
	m_AgProc = 0;
	m_Sequencial = 1;
	m_Box        = 0;
	m_DocBox     = 0;
	m_IsToInit   = false;
	
	// initialisation de la structure DCB
    m_DCB.dcbSize = sizeof(vipsDCB); // obligatoire, à remplir
    m_DCB.BaudRate= 0;               // défaut
    m_DCB.ByteSize= 0;               // défaut
    m_DCB.Parity= 0;                 // défaut
    m_DCB.StopBits= 0;               // défaut
    m_DCB.vipsTimeout= 0;            // défaut
    m_DCB.EvtChar= 0;                // défaut
    m_DCB.PortAddres= 0;             // défaut
    m_DCB.PortIRQ= 0;                // défaut
    m_DCB.Type= vipsLA93;            // obligatoire, à remplir
    m_DCB.Port = (unsigned char)m_ComPort;          // com par défaut

     // on définit le matériel composant le lecteur
    m_HD[0].hwType = vipsBoxes;           // escaninhos
    m_HD[0].hwState = m_Boxes;            // numero de escaninhos
    m_HD[1].hwType = vipsLenFeelers;      // longueur du seuil pour la détection des documents en double
    m_HD[1].hwState = 50;                 // valeur de cette longueur
    m_HD[2].hwType = vipsFeelers;         // seuil de détection d'un double
    m_HD[2].hwState = 35;                 // valeur de ce seuil
	m_HD[3].hwType = vipsPrinter;         // Impressora
	m_HD[3].hwState = 0;                  // 1 = Impressora ativada | 0 = desativada

    // on ouvre la communication avec le lecteur
    iRet = vipsOpen(&m_DCB);
    if (iRet != VER_NOERROR)  // si problème d'initialisation de la DLL vipsdrv
    {
        vipsRecupereStatusLibelle(m_strMsg, &iRet);
        TratamentoErro();
		return iRet;
    }

    GravaLog("Init - vipsOpen\n");
	
	// on affecte le matériel au système
    iRet = vipsSetHardware(4, m_HD); 
    if (iRet != VER_NOERROR)  // si problème de matériel
	{
        TratamentoErro();
		return iRet;
	}

	GravaLog("Init - vipsSetHardware\n");

	return 0;

}

void Done( void )
{
	if( m_fData )
	{
		fclose(m_fData);
		m_fData = NULL;
	}

	// fermeture de la com et libération de la DLL vipsdrv
    vipsClose();

	GravaLog("Done - vipsClose\n");

    // on libère la structure TVipsLC13
    VipsGrimLC13_Fin(&m_LC13);

	GravaLog("Done - VipsGrimLC13_Fin\n");

#ifdef _DEBUG
	if( m_fLog )
	{
		fclose(m_fLog);
		m_fLog = NULL;
	}
#endif
    
}

// tratamento de erro
void TratamentoErro( void )
{
	int iRet;

    // on éjecte les documents du lecteur
    iRet = vipsEject();

    sprintf(m_strLog, "TratamentoErro - vipsEject = %d\n", iRet);
	GravaLog(m_strLog);
}

// tratamento classico de erro 
long TratamentoClassico(TVipsLC13 *TVLC13, char *NomeArquivo, BOOL FacesNull)
{
	unsigned int Face1;     // type de la première face du document (RECTO, VERSO, ou FACE_NULLE)
	unsigned int Face2;     // type de la première face du document (RECTO, VERSO, ou FACE_NULLE)
    unsigned int NbImagesAEvacuer, i;
	int iRet;

	GravaLog("TratamentoClassico\n");

	if (FacesNull)
	{
		Face1 = FACE_NULLE;
		Face2 = FACE_NULLE;
	}
	else        // cas normal
	{
		Face1 = VERSO;
		Face2 = RECTO;
	}

	if (m_Reader == 1)
	{ 
		strcpy(m_UserDataBuffer.strCMC7, m_strCMC7);
		m_UserDataBuffer.strCodBar[0] = '\0';

		sprintf(m_strLog, "TratamentoClassico - CMC7 = %s\n", m_UserDataBuffer.strCMC7);
		GravaLog(m_strLog);
	}
	else
	{
		if (m_Reader == 2)
		{
			m_UserDataBuffer.strCMC7[0] = '\0';
			strcpy(m_UserDataBuffer.strCodBar, m_strCodBar);

			sprintf(m_strLog, "TratamentoClassico - CodBar = %s\n", m_UserDataBuffer.strCodBar);
			GravaLog(m_strLog);
		
		}
		else
		{
			strcpy(m_UserDataBuffer.strCMC7, m_strCMC7);
			strcpy(m_UserDataBuffer.strCodBar, m_strCodBar);

			sprintf(m_strLog, "TratamentoClassico - CMC7 = %s\n", m_UserDataBuffer.strCMC7);
			GravaLog(m_strLog);
			sprintf(m_strLog, "TratamentoClassico - CodBar = %s\n", m_UserDataBuffer.strCodBar);
			GravaLog(m_strLog);
		}
	}


	if( Face1 != FACE_NULLE )
	{
		m_Busy++;
		sprintf(m_strLog, "m_Busy = %d\n", m_Busy);
		GravaLog(m_strLog);
	}

	// envoi d'un ordre fichier à la tâche de compression
	iRet = VipsGrimLC13_OrdreFichier(TVLC13, Face1, NomeArquivo, &m_UserDataBuffer);
	if (iRet)
	{
		//TratamentoErro();

		return iRet;
	}

	GravaLog("TratamentoClassico - VipsGrimLC13_OrdreFichier Face1\n");

	// envoi d'un ordre fichier à la tâche de compression
	iRet = VipsGrimLC13_OrdreFichier(TVLC13, Face2, NomeArquivo, &m_UserDataBuffer);
	if (iRet)
	{
		//TratamentoErro();

		return iRet;
	}

	GravaLog("TratamentoClassico - VipsGrimLC13_OrdreFichier Face2\n");

	// on signale 2 ordres fichier envoyés
	iRet = VipsGrimLC13_SignaleNFichiersPresentsDansBuffer(TVLC13, 2);
	if (iRet)
	{
		// TratamentoErro();

		sprintf(m_strLog, "TratamentoClassico - VipsGrimLC13_SignaleNFichiersPresentsDansBuffer - iRet = %d\n", iRet);
		GravaLog(m_strLog);

		return iRet;
	}

	GravaLog("TratamentoClassico - VipsGrimLC13_SignaleNFichiersPresentsDansBuffer\n");
	
	// on attend deux buffers dma libres avant de dépiler
	iRet = VipsGrimLC13_AttenteNBuffersDmaLibres(TVLC13, 2);
	if (iRet)
	{

		sprintf(m_strLog, "TratamentoClassico - VipsGrimLC13_AttenteNBuffersDmaLibres - iRet = %d\n", iRet);
		GravaLog(m_strLog);

		// on teste le statut de VipsGrim
		iRet = VipsGrimLC13_GetStatus(TVLC13, m_NomeFuncaoRetornouErro, m_NomeFuncao, &m_iErr);
		if (iRet)
		{
			// TratamentoErro();

			sprintf(m_strLog, "TratamentoClassico - VipsGrimLC13_GetStatus - iRet = %d\n", iRet);
			GravaLog(m_strLog);
			
			return iRet;
		}

		GravaLog("TratamentoClassico - VipsGrimLC13_GetStatus\n");

		// on récupère le nombre d'images à évacuer
		NbImagesAEvacuer = TVLC13->TVBF.NbElemsDansBuffer;

		// suivant le cas
		for(i=0; i<NbImagesAEvacuer; i++)
		{
			iRet = VipsGrimLC13_GenereImageFictive(TVLC13);
			if (iRet)
			{
				// TratamentoErro();

				sprintf(m_strLog, "TratamentoClassico - VipsGrimLC13_GenereImageFictive - iRet = %d\n", iRet);
				GravaLog(m_strLog);

				return iRet;
			}

			GravaLog("TratamentoClassico - VipsGrimLC13_GenereImageFictive\n");

			iRet = VipsGrimLC13_SignaleNImagesPresentesDansBuffer(TVLC13, 1);
			if (iRet)
			{
				// TratamentoErro();

				sprintf(m_strLog, "TratamentoClassico - VipsGrimLC13_SignaleNImagesPresentesDansBuffer - iRet = %d\n", iRet);
				GravaLog(m_strLog);

				return iRet;
			}

			GravaLog("TratamentoClassico - VipsGrimLC13_SignaleNImagesPresentesDansBuffer\n");
		}

	}

	GravaLog("TratamentoClassico - VipsGrimLC13_AttenteNBuffersDmaLibres\n");
  
	// on attend un buffer image libre avant de dépiler
	do
	{        
		iRet = VipsGrimLC13_AttenteNBuffersImagesLibres(TVLC13, 1);

		sprintf(m_strLog, "TratamentoClassico - VipsGrimLC13_AttenteNBuffersImagesLibres - iRet = %d\n", iRet);
		GravaLog(m_strLog);
	}
	while (iRet);       // on boucle si la compression-écriture est trop longue

	// on attend un buffer image libre avant de dépiler
	do
	{
		iRet = VipsGrimLC13_AttenteNBuffersImagesLibres(TVLC13, 1);

		sprintf(m_strLog, "TratamentoClassico - VipsGrimLC13_AttenteNBuffersImagesLibres - iRet = %d\n", iRet);
		GravaLog(m_strLog);
	}
	while (iRet);       // on boucle si la compression-écriture est trop longue

	// on teste le statut de VipsGrim
	iRet = VipsGrimLC13_GetStatus(TVLC13, m_NomeFuncaoRetornouErro, m_NomeFuncao, &m_iErr);
	if (iRet)
	{
		// TratamentoErro();

		sprintf(m_strLog, "TratamentoClassico - VipsGrimLC13_GetStatus - iRet = %d\n", iRet);
		GravaLog(m_strLog);

		return iRet;
	}

	GravaLog("TratamentoClassico - VipsGrimLC13_GetStatus\n");

	return 0;
}

// tratamento de erro que libera os buffers
long TratamentoLimpaBuffer( void )
{
    unsigned int NbImagesATraiter, i;
	int iRet;

    GravaLog("TratamentoLimpaBuffer\n");
	
	// on attend que le buffers fichier se vide
    while (m_LC13.TVBF.NbElemsDansBuffer != 0)
        Sleep(100);

    // il faut s'assurer que toutes les images éventuellement capturées
    // soient dans le buffer image
    // très important !
    Sleep(2500);

    // on renseigne la structure TVipsUserDataBufferFichiers
	if (m_Reader == 1)
	{ 
		strcpy(m_UserDataBuffer.strCMC7, m_strCMC7);
		m_UserDataBuffer.strCodBar[0] = '\0';

		sprintf(m_strLog, "TratamentoLimpaBuffer - CMC7 = %s\n", m_UserDataBuffer.strCMC7);
		GravaLog(m_strLog);
	}
	else
	{
		if (m_Reader == 2)
		{
			m_UserDataBuffer.strCMC7[0] = '\0';
			strcpy(m_UserDataBuffer.strCodBar, m_strCodBar);

			sprintf(m_strLog, "TratamentoLimpaBuffer - CodBar = %s\n", m_UserDataBuffer.strCodBar);
			GravaLog(m_strLog);
		}
		else
		{
			strcpy(m_UserDataBuffer.strCMC7, m_strCMC7);
			strcpy(m_UserDataBuffer.strCodBar, m_strCodBar);

			sprintf(m_strLog, "TratamentoLimpaBuffer - CMC7 = %s\n", m_UserDataBuffer.strCMC7);
			GravaLog(m_strLog);
			sprintf(m_strLog, "TratamentoLimpaBuffer - CodBar = %s\n", m_UserDataBuffer.strCodBar);
			GravaLog(m_strLog);
		}
	}

    // on récupère le nombre d'images à traiter
    NbImagesATraiter = m_LC13.TVBI.NbElemsDansBuffer;

    // suivant le cas
    for(i=0; i<NbImagesATraiter; i++)
    {
        iRet = VipsGrimLC13_OrdreFichier(&m_LC13, FACE_NULLE, "", &m_UserDataBuffer);
        if (iRet)
		{
            TratamentoErro();
			return iRet;
		}

		GravaLog("TratamentoLimpaBuffer - VipsGrimLC13_OrdreFichier FaceNule\n");

        // on signale 1 ordre fichier envoyé
        iRet = VipsGrimLC13_SignaleNFichiersPresentsDansBuffer(&m_LC13, 1);
        if (iRet)
		{
            TratamentoErro();
			return iRet;
		}

		GravaLog("TratamentoLimpaBuffer - VipsGrimLC13_SignaleNFichiersPresentsDansBuffer\n");

        // on attend un buffer dma libre avant de dépiler
        iRet = VipsGrimLC13_AttenteNBuffersDmaLibres(&m_LC13, 1);
        if (iRet)
		{
            TratamentoErro();
			return iRet;
		}
        
		GravaLog("TratamentoLimpaBuffer - VipsGrimLC13_AttenteNBuffersDmaLibres\n");

        // on attend un buffer image libre avant de dépiler
        iRet = VipsGrimLC13_AttenteNBuffersImagesLibres(&m_LC13, 1);
        if (iRet)
		{
            TratamentoErro();
			return iRet;
		}

		GravaLog("TratamentoLimpaBuffer - VipsGrimLC13_AttenteNBuffersImagesLibres\n");
    }

    // on teste le statut de VipsGrim
    iRet = VipsGrimLC13_GetStatus(&m_LC13, m_NomeFuncaoRetornouErro, m_NomeFuncao, &m_iErr);
    if (iRet)
	{
        TratamentoErro();
		return iRet;
	}

	GravaLog("TratamentoLimpaBuffer - VipsGrimLC13_GetStatus\n");

    // on attend que le buffers fichier se vide
    while (m_LC13.TVBF.NbElemsDansBuffer != 0)
        Sleep(100);

	return 0;
}

// loop principal
long ProcessoCaptura( void )
{
    char NomeArquivo[256];
    int iRet, Count;
	MSG Msg;

	memset(&Msg, '\0', sizeof(MSG));
	Count = 0;

	GravaLog("ProcessoCaptura\n");

	// on règle la priorité de la tâche du programme
    // ainsi que celle du main
    VipsGrimLC13_ReglePrioriteProgrammeEtMain();

	GravaLog("ProcessoCaptura - VipsGrimLC13_ReglePrioriteProgrammeEtMain\n");

	GravaLog("ProcessoCaptura - Antes - vipsAddStation\n");

    vipsAddStation("O.B");

	GravaLog("ProcessoCaptura - vipsAddStation\n");

    while (true)
    {
		GravaLog("ProcessoCaptura - While(true)\n");

        iRet = VipsGrimLC13_GetStatus(&m_LC13, m_NomeFuncaoRetornouErro, m_NomeFuncao, &m_iErr);
        if (iRet)
		{
            TratamentoErro();
			return iRet;
		}

		GravaLog("ProcessoCaptura - VipsGrimLC13_GetStatus\n");

		// Tenta ler CMC7 e cod. de barras
		m_strCMC7[0] = '\0';
		m_strCodBar[0] = '\0';

		if (m_Reader == 1)
		{ 
			iRet = vipsReadCMC7(m_strCMC7, 63);

			sprintf(m_strLog, "ProcessoCaptura - vipsReadCMC7 = %s\n", m_strCMC7);
			GravaLog(m_strLog);
		}
		else
		{
			if (m_Reader == 2)
			{
				iRet = vipsReadSecond(m_strCodBar, 63);

				sprintf(m_strLog, "ProcessoCaptura - vipsReadSecond = %s\n", m_strCodBar);
				GravaLog(m_strLog);
			}
			else
			{
				iRet = vipsReadDouble(m_strCMC7, 63, m_strCodBar, 63);

				sprintf(m_strLog, "ProcessoCaptura - vipsReadDouble = %s e %s\n", m_strCMC7, m_strCodBar);
				GravaLog(m_strLog);
			}
		}
		
        //sprintf(NomeArquivo, "%04.4ld%05.5ld%05.5ld", m_AgProc, m_Lote, m_Sequencial);
		sprintf(NomeArquivo, "%09.9ld%05.5ld", m_Lote, m_Sequencial);

		sprintf(m_strLog, "ProcessoCaptura - NomeArquivo = %s\n", NomeArquivo);
		GravaLog(m_strLog);

		m_Sequencial++;
		m_DocBox++;
		if( m_DocBox > m_MaxDocBox )
		{
			m_Box++;
			if( m_Box > m_Boxes - 1 )
			{
				m_Box = 1;
			}
			m_DocBox = 1;
		}

		sprintf(m_strLog, "ProcessoCaptura - m_Box = %d, m_DocBox = %ld\n", m_Box, m_DocBox);
		GravaLog(m_strLog);

		sprintf(m_strLog, "ProcessoCaptura - m_BoxDefault = %ld\n", m_BoxDefault);
		GravaLog(m_strLog);

		// Seta o escaninho 
		if( m_BoxDefault > 0 )
			iRet = vipsSortDoc(m_BoxDefault, "", "", 0);
		else
			// iRet = vipsSortDoc(m_Box, "", "", 0);
			// Alteracao da ordem dos escaninhos 
			iRet = vipsSortDoc(m_aBoxes[m_Box], "", "", 0);

		sprintf(m_strLog, "ProcessoCaptura - vipsSortDoc = %d\n", iRet);
		GravaLog(m_strLog);

		// acabou os documentos do alimentador
		if( iRet == -2 )
		{
			GravaLog("ProcessoCaptura - iRet == -2\n");

			break;
		}
		else if( iRet == -1 || iRet == -51 || iRet == -55 || iRet == -59 ||
			     iRet == -61 )
		{
			sprintf(m_strLog, "ProcessoCaptura - iRet = %d\n", iRet);
			GravaLog(m_strLog);

			vipsRecupereStatusLibelle(m_strMsg, &iRet);

			sprintf(m_strLog, "ProcessoCaptura - vipsRecupereStatusLibelle = iRet = %d\n", iRet);
			GravaLog(m_strLog);

			TratamentoLimpaBuffer();
			return iRet;
		}
		else if( iRet == -62 )
		{
			sprintf(m_strLog, "ProcessoCaptura - iRet = %d\n", iRet);
			GravaLog(m_strLog);

			vipsRecupereStatusLibelle(m_strMsg, &iRet);

			sprintf(m_strLog, "ProcessoCaptura - vipsRecupereStatusLibelle = iRet = %d\n", iRet);
			GravaLog(m_strLog);

			vipsSetHardware(1, m_HD); 

			strcpy(m_strLog, "ProcessoCaptura - vipsSetHardware\n");
			GravaLog(m_strLog);
			
			TratamentoLimpaBuffer();
			return iRet;
		}
		else if( iRet == -50 )
		{
			GravaLog("ProcessoCaptura - iRet == -50\n");

			vipsRecupereStatusLibelle(m_strMsg, &iRet);

			sprintf(m_strLog, "ProcessoCaptura - vipsRecupereStatusLibelle = iRet = %d\n", iRet);
			GravaLog(m_strLog);

			TratamentoErro();
			TratamentoClassico(&m_LC13, NomeArquivo, TRUE);

			return -50;
		}
		else if( iRet < 0 )
		{
			sprintf(m_strLog, "ProcessoCaptura - iRet = %d\n", iRet);
			GravaLog(m_strLog);

			vipsRecupereStatusLibelle(m_strMsg, &iRet);

			sprintf(m_strLog, "ProcessoCaptura - vipsRecupereStatusLibelle = iRet = %d\n", iRet);
			GravaLog(m_strLog);

			TratamentoErro();
			TratamentoLimpaBuffer();
			return iRet;
		}
		
		iRet = TratamentoClassico(&m_LC13, NomeArquivo, FALSE);
		if (iRet)
		{
			if( iRet != 293 )
			{
				TratamentoErro();
				TratamentoLimpaBuffer();
			}
			return iRet;
		}
		if( Count++ % 10 == 0 )
		{
			PeekMessage(&Msg, NULL, 0, 0, PM_NOREMOVE);
		}

    }

    // on attend que le buffers fichier se vide
    while (m_LC13.TVBF.NbElemsDansBuffer != 0)
        Sleep(100);

	return 0;
}

// callback responsavel pela gravacao da imagem
int __stdcall FonctionUserDataBufferFichiers(char MomentAppel, unsigned int InformationFace, char *NomFichier,
    TVipsBMP *TVB, BOOL *Ecriture, USERDATA_BUFFER_FICHIERS *PUserDataBufferFichiers)

{
	char strNomeArq[256];
	char strNomeCompleto[512];
	TRetorno Retorno;
	TVipsUserDataBufferFichiers *pUserDataBuffer = (TVipsUserDataBufferFichiers *)PUserDataBufferFichiers;
	size_t Gravado;
    int iRet;

	GravaLog("CallBack\n");

	switch (MomentAppel)                // quel moment est-ce ?
	{
		case USERDATA_AVANT_ECRITURE :  

			//if (EXPANSION_DYNAMIC && InformationFace == RECTO)
			if (EXPANSION_DYNAMIC)
			{
			   BMP_ExpansionDynamique(TVB, SEUIL_EXPANSION);

			   GravaLog("CallBack - BMP_ExpansionDynamique\n");
			}

			// réhausse le contraste d'une image par le biais d'un filtre Laplacien
			//if (REHAUSSEMENT_CONTRASTE && InformationFace == RECTO)
			if (REHAUSSEMENT_CONTRASTE)
			{
				iRet = BMP_RehaussementContraste(TVB, LAMBDA);
				if (iRet)
				{
					TratamentoErro();
					m_Busy--;
					return iRet;
				}

				GravaLog("CallBack - BMP_RehaussementContraste\n");
			} 

			if( m_Threshold > 0 )
			{
				//if( InformationFace == RECTO )
				iRet = BMP_CoupeBords(TVB, &m_BMPOut, (unsigned char)m_Threshold, FOND_NOIR );
				//else
				//	iRet = BMP_CoupeBords(TVB, &m_BMPOut, (unsigned char)(m_Threshold+25), FOND_NOIR );

				if (iRet)
				{
					TratamentoErro();
					m_Busy--;
					return iRet;
				}

				GravaLog("CallBack - BMP_CoupeBords\n");
			}
			else
			{
				memcpy(&m_BMPOut, TVB, sizeof(TVipsBMP));

				GravaLog("CallBack - Copia BMP\n");
			}
      
			if (InformationFace == RECTO)
			{
				sprintf(m_strLog, "CallBack - Altura da Imagem = %ld\n", m_BMPOut.BmI->bmiHeader.biHeight);
				GravaLog(m_strLog);

				// se imagem esta em 100 dpi e tem menos de 240 pixels de altura (6 cm)
				if( m_Resolution == RES100DPI && m_BMPOut.BmI->bmiHeader.biHeight <= 240 )
				{
					m_IsToInit = true;
					GravaLog("CallBack - Foi detectado documento com altura insuficiente\n");
				}
				// se imagem esta em 200 dpi e tem menos de 480 pixels de altura (6 cm)
				else if( m_Resolution == RES200DPI && m_BMPOut.BmI->bmiHeader.biHeight <= 480 )
				{
					m_IsToInit = true;
					GravaLog("CallBack - Foi detectado documento com altura insuficiente\n");
				}
			}

			switch( m_ImageType )
			{
				case 1: // BMP
					strcpy(strNomeArq, NomFichier);
					if (InformationFace == RECTO)
						strcat(strNomeArq, "f");
					else
						strcat(strNomeArq, "v");
					strcat(strNomeArq, ".bmp");

					sprintf(strNomeCompleto, "%s%09.9ld\\%s", m_Directory, m_Lote, strNomeArq);

					iRet = BMP_Ecrit(&m_BMPOut, strNomeCompleto, TRUE);

					sprintf(m_strLog, "CallBack - Grava BMP = %s\n", strNomeArq);
					GravaLog(m_strLog);

					break;
				case 2: // TIFF

/*                  Problemas na gravacao como Tiff, a funcao sempre
                    retorna 181. No Arquivo vipsimage.h este valor
					refere-se a seguinte constante: IMAGES_DLL_NON_TROUVEE_ERR

					strcpy(strNomeArq, NomFichier);
					if (InformationFace == RECTO)
						strcat(strNomeArq, "f");
					else
						strcat(strNomeArq, "v");
					strcat(strNomeArq, ".tif");

					sprintf(strNomeCompleto, "%s%09.9ld\\%s", m_Directory, m_Lote, strNomeArq);

					iRet = BMP_BMP2TIFFG4_Disque(&m_BMPOut, strNomeCompleto, 128, 100, FALSE);

					sprintf(m_strLog, "CallBack - Grava TIFF = %s\n", strNomeArq);
					GravaLog(m_strLog);

					break; */
				case 3: // JPG
					strcpy(strNomeArq, NomFichier);
					if (InformationFace == RECTO)
						strcat(strNomeArq, "f");
					else
						strcat(strNomeArq, "v");
					strcat(strNomeArq, ".jpg");

					sprintf(strNomeCompleto, "%s%09.9ld\\%s", m_Directory, m_Lote, strNomeArq);

					iRet = BMP_BMP2JPEG_DisqueIJG(&m_BMPOut, strNomeCompleto, m_Compress);

					sprintf(m_strLog, "CallBack - Grava JPG = %s\n", strNomeArq);
					GravaLog(m_strLog);
					
					break;
			}
			
			
			if (iRet)
			{
				TratamentoErro();
				BMP_Fin(&m_BMPOut);

				sprintf(m_strLog, "CallBack - Gravacao da imagem falhou. Erro = %d\n", iRet);
				GravaLog(m_strLog);

				m_Busy--;
				return iRet;
			}

			GravaLog("CallBack - Gravacao da imagem concluida\n");

			// you must test return code here
			iRet = BMP_Fin(&m_BMPOut);
			if (iRet)
			{
				TratamentoErro();
				m_Busy--;
				return iRet;
			}

			GravaLog("CallBack - BMP_Fin\n");

			if (InformationFace == RECTO)
			{
				memset(&Retorno, ' ', sizeof(TRetorno));

				Retorno.Tipo[0] = 'A';
				Retorno.Origem[0] = '0';
				Retorno.CrLf[0] = 0x0D;
				Retorno.CrLf[1] = 0x0A;

				memcpy(Retorno.Frente, NomFichier, 14);
				Retorno.Frente[14] = 'f';
				memcpy(Retorno.Verso, NomFichier, 14);
				Retorno.Verso[14] = 'v';
				
				switch( m_ImageType )
				{
					case 1:
						memcpy(&Retorno.Frente[15], ".bmp", 4);
						memcpy(&Retorno.Verso[15], ".bmp", 4);
						break;
					case 2:
						memcpy(&Retorno.Frente[15], ".tif", 4);
						memcpy(&Retorno.Verso[15], ".tif", 4);
						break;
					case 3:
						memcpy(&Retorno.Frente[15], ".jpg", 4);
						memcpy(&Retorno.Verso[15], ".jpg", 4);
						break;
				}

				if( (strlen( pUserDataBuffer->strCodBar ) > 0) &&
				     !( strlen( pUserDataBuffer->strCodBar ) == 1 && pUserDataBuffer->strCodBar[0] ==  '!' ) )
				{
					Retorno.Tipo[0] = 'B';
					memcpy(Retorno.Leitura, pUserDataBuffer->strCodBar, strlen( pUserDataBuffer->strCodBar ));
				}
				else if( (strlen( pUserDataBuffer->strCMC7 ) > 0) && 
					 !( strlen( pUserDataBuffer->strCMC7 ) == 1 && pUserDataBuffer->strCMC7[0] == '!' ) )
				{
					memcpy(Retorno.Leitura, pUserDataBuffer->strCMC7, strlen( pUserDataBuffer->strCMC7 ));
				}
				
				sprintf(m_strLog, "CallBack - strCodBar = %63.63s\n", pUserDataBuffer->strCodBar);
				GravaLog(m_strLog);
				sprintf(m_strLog, "CallBack - strCMC7   = %63.63s\n", pUserDataBuffer->strCMC7);
				GravaLog(m_strLog);

				sprintf(m_strLog, "CallBack - Retorno.Leitura = %63.63s\n", Retorno.Leitura);
				GravaLog(m_strLog);
				sprintf(m_strLog, "CallBack - Retorno.Frente = %19.19s\n", Retorno.Frente);
				GravaLog(m_strLog);
				sprintf(m_strLog, "CallBack - Retorno.Verso = %19.19s\n", Retorno.Verso);
				GravaLog(m_strLog);
				
				if( m_fData )
				{
					Gravado = fwrite(&Retorno, 1, sizeof(TRetorno), m_fData);
					if( Gravado != sizeof(TRetorno) )
					{
						m_Busy--;
						// erro na gravacao do registro
						return -1001;
					}

				}
				else
				{
					m_Busy--;
					// arquivo de dados nao esta aberto
					return -1002;
				}

				GravaLog("CallBack - Gravacao dos dados concluida\n");
			}

			*Ecriture = FALSE;

			if (InformationFace == RECTO)
			{
				m_Busy--;
				sprintf(m_strLog, "m_Busy = %d\n", m_Busy);
				GravaLog(m_strLog);
			}

			break;

		default :  
			break;

	}

    return 0;       // tout est ok
}

// grava log
void GravaLog( char *Log )
{
#ifdef _DEBUG
	if( m_fLog )
	{
		fputs(Log, m_fLog);
		fflush(m_fLog);
	}
#endif
}

