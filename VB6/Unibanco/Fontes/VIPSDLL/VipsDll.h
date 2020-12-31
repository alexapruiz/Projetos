
#include "VipsDrv.h"
#include "VipsGrim.h"

#define VIDEO_CARD                          CARTE_DIPIX_LPG
#define NB_BUFFERS_IMAGES                   6
#define NB_BUFFERS_FILES                    6
#define NB_BUFFERS_DMA                      2
#define CORRECTION_LUMINEUSE                FALSE
#define ETALON_FILE                         "etalon.eta"
#define EXPANSION_DYNAMIC                   TRUE
#define SEUIL_EXPANSION                     0
#define REHAUSSEMENT_CONTRASTE              TRUE
#define LAMBDA                              0.1
#define MAX_BOXES                           21
#define MAX_DOC_BOX                         120

long m_ComPort;                    // Serial 1
long m_Boxes;                      // Nro. de Escanhinhos
char m_Resolution;                 // Resolucao
long m_Reader;                     // Ambas
long m_ImageType;                  // JPG
long m_Compress;                   // Fator de Compressao
long m_Threshold;                  // Fator para corte das imagens
char m_Directory[256];             // Diretorio das imagens
char m_CameraFile[256];            // Path complento do arq. conf. da camera
long m_Sequencial;                 // Seq. da imagem dentro do lote
long m_AgProc;                     // Ag. Processadora
long m_Lote;                       // Lote 
short m_Box;                       // Escaninho atualmente usado
short m_BoxDefault;                // Se selecionado, sempre usara este
long  m_DocBox;                    // Qtde de documentos no escaninho atual
long  m_MaxDocBox;                 // Qtde Maxima de docs por escaninho

// Alteracao da ordem dos escaninhos 
short m_aBoxes[MAX_BOXES+1];       // Array para ordenar os escaninhos

vipsDCB m_DCB;          // structure de contrôle de la série pour vipsdrv
vipsHARDWARE m_HD[10];  // structure définissant le matériel utilisé par le lecteur
TVipsBMP m_BMPOut;      // Usado para converter a imagem
TVipsLC13 m_LC13;       // structure TVipsLC13, pour la gestion des cartes vidéo et des images

char m_strCMC7[64];     // string de CMC7
char m_strCodBar[64];   // string de Codigo de Barras
char m_strMsg[256];     // String de Mensagens
char m_strLog[256];

typedef struct
{
    char strCMC7[64];
	char strCodBar[64];
} TVipsUserDataBufferFichiers;

TVipsUserDataBufferFichiers m_UserDataBuffer;

typedef struct
{
    char Tipo[1],
         Leitura[63],
         Frente[19],
         Verso[19],
         Origem[1],
		 CrLf[2];
} TRetorno;

char m_NomeFuncaoRetornouErro[256];
char m_NomeFuncao[256];
unsigned int m_iErr;

FILE *m_fData;            // Handle do arquivo de dados
FILE *m_fLog;

int m_Busy;              // Sinaliza se Callback executanto
BOOL m_IsToInit;          // Sinaliza se existe a necessidade de Inicializar a VIPS

// tratamento de erro
void TratamentoErro( void );

PTR_FonctionUserDataBufferFichiers_STDCALL pFunctionFiler_STDCALL;

// callback responsavel pela gravacao da imagem
int __stdcall FonctionUserDataBufferFichiers(char MomentAppel, unsigned int InformationFace, char *NomFichier,
    TVipsBMP *TVB, BOOL *Ecriture, USERDATA_BUFFER_FICHIERS *PUserDataBufferFichiers);

// tratamento classico de erro 
long TratamentoClassico(TVipsLC13 *TVLC13, char *NomeArquivo, BOOL FacesNull);

// tratamento de erro que libera os buffers
long TratamentoLimpaBuffer( void );

long Init( void );

void Done( void );

// loop principal
long ProcessoCaptura( void );

// grava log
void GravaLog( char *Log );

// Inicializa Dll's 
long WINAPI VIPS_Init( void );

// Libera as Dll's
void WINAPI VIPS_Done( void );

// Seleciona a porta serial
void WINAPI VIPS_SetComPort( long ComPort );

// Seleciona a resolucao ( 100 ou 200 DPI )
void WINAPI VIPS_SetResolution( long Res );

// Seleciona leitora ( 1 = CMC7; 2 = Cod. Barras; 3 = Ambas )
void WINAPI VIPS_SetReader( long Reader );

// Seleciona a quantidade de Escaninhos ( Nro. de escaninhos deve ser impar )
void WINAPI VIPS_SetBoxes( long Boxes );

// Seleciona a quantidade maxima de docs por escanhinho
void WINAPI VIPS_SetMaxDocBox( long MaxDocBox );

// Seleciona o escanhinho default ( Default = 0 : Nenhum )
void WINAPI VIPS_SetBoxDefault( long BoxDefault );

// Seleciona o tipo de imagem gerada ( 1 = BMP; 2 = TIF; 3 = JPG )
void WINAPI VIPS_SetImageType( long ImageType );

// Seleciona o fator de compressao do JPG ( defualt = 30 ) 
void WINAPI VIPS_SetCompress( long Fator );

// Seleciona se deve cortar as bordas 
// ( Valor-> 0 a 255, Zero nao corta, default = 50. Quanto mais alto, mais corta )
void WINAPI VIPS_SetCutBords( long Valor );

// Seleciona o diretorio onde serao gravadas as imagens
void WINAPI VIPS_SetImageDirectory( char *Diretorio );

// Seleciona o arq de conf. da camera ( nao precisa do caminho )
void WINAPI VIPS_SetCameraFile( char *CameraFile );

// Executa a captura
long WINAPI VIPS_Captura( long AgProc, long Lote, char *ArqRetorno, long Append );

// Executa a captura
long WINAPI VIPS_Recaptura( long AgProc, long Lote, long SeqInicial, char *ArqRetorno, long Append );

// Reseta as Dll's
long WINAPI VIPS_Reset( void );
