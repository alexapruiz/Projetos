#ifndef __VIPS_H__
#define __VIPS_H__

int m_iRet;               // resultado das funcoes

vipsDCB m_DCB;          // structure de contrôle de la série pour vipsdrv
vipsHARDWARE m_HD[10];  // structure définissant le matériel utilisé par le lecteur
TVipsBMP m_BMPOut;      // Usado para converter a imagem
TVipsLC13 m_LC13;       // structure TVipsLC13, pour la gestion des cartes vidéo et des images

char m_strCMC7[64];     // string de CMC7
char m_strCodBar[64];   // string de Codigo de Barras
char m_strMsg[256];     // String de Mensagens

typedef struct
{
    char strCMC7[64];
	char strCodBar[64];
} TVipsUserDataBufferFichiers;

TVipsUserDataBufferFichiers m_UserDataBuffer;

char m_NomeFuncaoRetornouErro[256];
char m_NomeFuncao[256];
unsigned int m_iErr;

char m_NomeArqFrente[512];
char m_NomeArqVerso[512];
char m_NomeArqDados[256];

FILE *m_fData;            // Handle do arquivo de dados

// callback responsavel pela gravacao da imagem
int __stdcall FunctionFiler(char MomentAppel, unsigned int InformationFace, char *NomFichier,
    TVipsBMP *TVB, BOOL *Ecriture, USERDATA_BUFFER_FICHIERS *PUserDataBufferFichiers);

// tratamento de erro
void TratamentoErro(TVipsLC13 *TVLC13, char TypeError, char *NomeFuncaoRetornouErro, char *NomeFuncao, int Err);

// tratamento classico de erro 
void TratamentoClassico(TVipsLC13 *TVLC13, char *NomeArqVerso, char *NomeArqFrente, BOOL FacesNull);

// tratamento de erro contornavel
void TratamentoErroContornavel(BOOL IncrementationDoc);




#endif /* __VIPS_H__ */
