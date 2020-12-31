#ifndef __VIPSGRAB_H__
#define __VIPSGRAB_H__

#ifdef __cplusplus
extern "C" {
#endif

/******************************************************************/
/*                             LIBRAIRIE                          */
/*----------------------------------------------------------------*/
/* Librairie : vipsgrab.dll                                       */
/* Titre : Gestion des cartes d'acquisitions MAGIC et             */
/*         PULSAR                                                 */
/* Contenu: - initialisation des cartes                           */
/*          - d�marrage de la capture                             */
/*          - acquisition d'une image                             */
/*          - recup�ration de l'image captur�e                    */
/*          - arr�t de la capture                                 */
/*          - choix du mode d'affichage vid�o (MAGIC seulement)   */
/* Version : v1.20                                                */
/* D�velopp�e par : M. Tola                                       */
/*----------------------------------------------------------------*/
/* MODIFICATIONS                                                  */
/*----------------------------------------------------------------*/
/* Date de la modification : 11/12/97                             */
/* Code : - LEX                                                   */
/*        - gestion de la carte PULSAR                            */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 22/12/97                             */
/* Code : - LEX                                                   */
/*        - modularit� de la DLL au niveau des cartes utilis�es   */
/*        - renommage de la DLL en VipsGrab.dll                   */
/*        - renomage des fonctions GetDLLVideoVersion en          */
/*          GetDLLVipsGrabVersion et Video_Erreur en              */
/*          VipsGrab_Erreur                                       */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 16/01/98                             */
/* Code : - LEX                                                   */
/*        - gestion 100 et 200 DPI                                */
/*        - gestion fen�trage de la capture                       */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 20/03/98                             */
/* Code : - LEX                                                   */
/*        - chargement dynamique des dlls                         */
/*        - acquisition asynchrone                                */
/*        - fonctions �v�nementielles                             */
/*        - cr�ation de la fct. Pulsar_Init2                      */
/*        - Pulsar_Init appelle d�sormais Carte_Pulsar_Init       */
/*        - MAGIC_DemarreCapture et Magic_FinCapture retournent   */
/*          d�sormais un int                                      */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 21/04/98                             */
/* Code : - LEX                                                   */
/*        - VipsGrab_DechargeDlls modifi�e pour mieux g�rer le    */
/*          d�chargement dynamique des Dlls                       */
/*        - adaptation de PULSAR_AttacheFonctionEvenementielle    */
/*          � Mil-Lite 5.1 (gestion des �v�nements GRAB_START     */
/*          et GRAB_END)                                          */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 29/04/98                             */
/* Code : - LEX                                                   */
/*        - gestion des allocations dynamiques des buffers DMA    */
/*          sur la carte PULSAR ou en zone non swappable          */
/*          (cf. Carte_Pulsar_Init)                               */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 26/08/98                             */
/* Code : - LDB                                                   */
/*        - correction positionnement fen�trage acquisition en    */
/*          100 DPI                                               */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/******************************************************************/

// includes de Windows
#include <stdio.h>
#include <conio.h>
#include <stdlib.h>
#include <string.h>
#include <windows.h>
#include <process.h>
#include <io.h>
#include <fcntl.h>
#include <sys\stat.h>

// nombre maximal de buffers DMA autoris�s
#define NB_MAX_BUFFERS_DMA            10

// taille maximale d'une chaine de caract�res
#define TAILLE_MAX_CHAINE             256

// "profondeur" d'un plan image,
// en 256 niveaux de gris, 1 pixel = 1 octet 
#define DATA_DEPTH                    8

// #define des erreurs de la DLL
#define VIDEO_BUFFER_DMA_PLEIN_ERR                221   // buffer DMA plein
#define VIDEO_TROP_DE_BUFFERS_DMA_DEMANDES_ERR    222   // trop de buffers DMA demand�s par rapport au nombre autoris�
#define VIDEO_BUFFER_DMA_VIDE_ERR                 223   // buffer DMA vide

// Pr�d�claration des fonctions export�es
void __stdcall GetDLLVipsGrabVersion(char *Version);   // retourne le num�ro de la version de la DLL sous forme d'une chaine de caract�res
int __stdcall VipsGrab_Erreur(unsigned int NumeroErreur, char *ChaineErreur);  // retourne la chaine de caract�res correspondante au num�ro d'erreur pass� en param�tre

/********************************************************/

// include de magicnt.dll : gestion des cartes MAGIC
#ifndef MSC_NT
    #define MSC_NT
#endif

#ifndef _MT
    #define _MT
#endif

#include <magic.h>

// plan image sur lequel on "travaille", en 256 niveaux de gris,
// il s'agit du rouge
#define SURFACE                       ML_IMAGE_RED

// #define des diff�rents modes d'affichages possibles
#define VGA                 11      // r�alise uniquement l'affichage VGA
#define IMAGE               12      // r�alise uniquement l'affichage du buffer image de la carte d'acquisition
#define VGA_ET_IMAGE        13      // superpose l'affichage VGA et le buffer image quand la couleur de fond est 0 donc noire

// #define des erreurs de la DLL
#define VIDEO_INIT_DEFAULT_ERR                201   // erreur retourn�e par mlInitDefault
#define VIDEO_CAM_LOAD_ERR                    202   // erreur retourn�e par mlCamLoad
#define VIDEO_DMA_BUFFER_ALLOC_ERR            203   // erreur retourn�e par mlDmaBufferAlloc
#define VIDEO_DMA_BUFFER_GET_PTR_ERR          204   // erreur retourn�e par mlDmaBufferGetPtr
#define VIDEO_DMA_BUFFER_SELECT_ERR           205   // erreur retourn�e par mlDmaBufferSelect
#define VIDEO_DMA_BUFFER_DESELECT_ERR         206   // erreur retourn�e par mlDmaBufferDeselect
#define VIDEO_DMA_BUFFER_FREE_ERR             207   // erreur retourn�e par mlDmaBufferFree
#define VIDEO_DMA_WAIT_ERR                    208   // erreur retourn�e par mlDmaWait
#define VIDEO_DMA_READ_TRANS_ERR              209   // erreur retourn�e par mlDmaReadTrans
#define VIDEO_DMA_BUFFER_READ_AREA_ERR        210   // erreur retourn�e par mlDmaBufferReadArea

#define VIDEO_CAM_USER_BIT_OUT_STATE_ERR      211   // erreur retourn�e par mlCamUserBitOutState
#define VIDEO_CAM_USER_BIT_OUT_ERR            212   // erreur retourn�e par mlCamUserBitOut

// cf. #define 221, 222, 223 ci-dessus
#define VIDEO_MODE_AFFICHAGE_INCONNU              224   // mode d'affichage inconnu

// d�finition de la structure TVipsMAGIC
// structure permettant la gestion des cartes MAGIC
typedef struct
{
    TCAMERA    Camera;      // structure contenant les param�tres de
                            // r�glages de la cam�ra

    char NomFichierDCF[TAILLE_MAX_CHAINE];      // nom du fichier DCF
                                                // contenant les param�tres de
                                                // r�glages de la cam�ra

    BOOL UserBitOut;        // bool�en permettant de g�n�rer un front montant
                            // (sur le UserBit 0 de la carte MAGIC) apr�s chaque acquisition dans MAGIC_Acquisition

    BOOL DoubleBuffer;      // bool�en permettant de g�rer le "double buffering"

    unsigned int LargeurImage, HauteurImage;    // largeur et hauteur d'une image captur�e (en pixels)
    
    unsigned int NbBuffersDma;          // nombre de buffers DMA utilis�s

    short IdentificateurBufferDma[NB_MAX_BUFFERS_DMA];      // tableau d'identificateurs de chaque buffer DMA,
                                                            // cet identificateur est utile pour les fonctions de transfert DMA                                                                    

    unsigned char *AdresseBufferDma[NB_MAX_BUFFERS_DMA];    // tableau de pointeurs de chaque buffer DMA

    unsigned int NbElemsDansBufferDma;      // nombre d'images dans le buffer DMA
    unsigned int PointeurInBufferDma;       // index d'entr�e dans le buffer DMA
    unsigned int PointeurOutBufferDma;      // index de sortie dans le buffer DMA

    // num�ro de page virtuelle (gestion du "double buffering")
    short Page;

    unsigned char        *future[5];  // tableau de pointeurs pour "r�server de la place"
                                      // pour des extensions futures et amener la taille
                                      // de la structure � 832 octets
    
} TVipsMAGIC;

// Pr�d�claration des fonctions export�es
int __stdcall MAGIC_Init(TVipsMAGIC *TVM,           // initialisation de la carte MAGIC
                         char *NomFichierDCF,
                         unsigned int NbBuffersDma,
                         unsigned int *XSize, unsigned int *YSize,
                         BOOL UserBitOut,
                         BOOL DoubleBuffer);
int __stdcall MAGIC_DemarreCapture(void);           // d�marre la capture
int __stdcall MAGIC_Acquisition(TVipsMAGIC *TVM);   // acquisition d'une image
int __stdcall MAGIC_RecupereImage(TVipsMAGIC *TVM, unsigned char *RawData); // recup�re une image captur�e
int __stdcall MAGIC_FinCapture(void);               // arr�te la capture
int __stdcall MAGIC_Fin(TVipsMAGIC *TVM);           // lib�re la carte MAGIC
int __stdcall MAGIC_Affichage(int ModeAffichage);   // change le mode d'affichage

/********************************************************/

// include de mil.dll et milpul.dll : gestion des cartes PULSAR
#include <mil.h>

#define RES100DPI      1            // capture en 100 dpi
#define RES200DPI      2            // capture en 200 dpi : n�cessite un fichier DCF ad�quat

#define VIDEO_MAPP_ALLOC_ERR        231     // erreur retourn�e par MappAlloc
#define VIDEO_MSYS_ALLOC_ERR        232     // erreur retourn�e par MsysAlloc
#define VIDEO_MDIG_ALLOC_ERR        233     // erreur retourn�e par MdigAlloc
#define VIDEO_MDIG_CONTROL_ERR      234     // erreur retourn�e par MdigControl
#define VIDEO_MDIG_CHANNEL_ERR      235     // erreur retourn�e par MdigChannel
#define VIDEO_BUF_ALLOC_2D_ERR      236     // erreur retourn�e par MbufAlloc2d
#define VIDEO_MBUF_FREE_ERR         237     // erreur retourn�e par MbufFree
#define VIDEO_MDIG_FREE_ERR         238     // erreur retourn�e par MdigFree
#define VIDEO_MSYS_FREE_ERR         239     // erreur retourn�e par MsysFree
#define VIDEO_MAPP_FREE_ERR         240     // erreur retourn�e par MappFree
#define VIDEO_MDIG_HALT_ERR         241     // erreur retourn�e par MdigHalt
#define VIDEO_MDIG_GRAB_ERR         242     // erreur retourn�e par MdigGrab
#define VIDEO_MDIG_GRAB_WAIT_ERR    243     // erreur retourn�e par MdigGrabWait
#define VIDEO_MBUF_GET_2D_ERR       244     // erreur retourn�e par MbufGet2d
#define VIDEO_MAPP_CONTROL_ERR      245     // erreur retourn�e par MappControl
#define VIDEO_MDIG_INQUIRE_ERR      246     // erreur retourn�e par MdigInquire

#define VIDEO_MDIG_GRAB_TIMEOUT_ERR 251     // erreur de TimeOut pendant la capture

#define VIDEO_LARGEUR_FENETRE_ERR    261    // erreur sur la largeur de la fen�tre
#define VIDEO_HAUTEUR_FENETRE_ERR    262    // erreur sur la hauteur de la fen�tre
#define VIDEO_LARGEUR_IMAGE_ERR      263    // erreur sur la largeur de l'image
#define VIDEO_HAUTEUR_IMAGE_ERR      264    // erreur sur la hauteur de l'image
#define VIDEO_OFFSETX_FENETRE_ERR    265    // erreur sur l'offset en X de la fen�tre
#define VIDEO_OFFSETY_FENETRE_ERR    266    // erreur sur l'offset en Y de la fen�tre
#define VIDEO_FENETRE_ERR            267    // erreur, mode de fen�trage non actif

// d�finition de la structure TVipsPULSAR
// structure permettant la gestion des cartes PULSAR
typedef struct
{
    char NomFichierDCF[TAILLE_MAX_CHAINE];      // nom du fichier DCF
                                                // contenant les param�tres de
                                                // r�glages de la cam�ra

    // non utilis� pour le moment
    BOOL UserBitOut;        // bool�en permettant de g�n�rer un front montant
                            // (sur le UserBit 0 de la carte PULSAR) apr�s chaque acquisition dans PULSAR_Acquisition

    // non utilis� pour le moment
    BOOL DoubleBuffer;      // bool�en permettant de g�rer le "double buffering"

    unsigned int LargeurImage, HauteurImage;    // largeur et hauteur d'une image captur�e (en pixels)
    
    unsigned int NbBuffersDma;          // nombre de buffers DMA utilis�s

    unsigned int NbElemsDansBufferDma;      // nombre d'images dans le buffer DMA
    unsigned int PointeurInBufferDma;       // index d'entr�e dans le buffer DMA
    unsigned int PointeurOutBufferDma;      // index de sortie dans le buffer DMA

    // non utilis� pour le moment
    short Page;     // num�ro de page virtuelle (gestion du "double buffering")

    unsigned char        *future[3];  // tableau de pointeurs pour "r�server de la place"
                                      // pour des extensions futures et amener la taille
                                      // de la structure � 384 octets

    /****************/

    unsigned int DureeTimeOut;      // TimeOut de la capture d'image

    MIL_ID MilApplication;   // identificateur de "l'application" MIL
    MIL_ID MilSysteme;       // identificateur du "syst�me" MIL
    MIL_ID MilCamera;        // identificateur de la "cam�ra" MIL

    MIL_ID IdentificateurBufferDma[NB_MAX_BUFFERS_DMA];     // identificateurs des "buffers DMA" MIL

    char Resolution;        // r�solution de la capture

    BOOL Fenetre;            // pour utiliser le fen�trage ou non
    unsigned int LargeurFenetre, HauteurFenetre;        // largeur et hauteur de la fen�tre
    unsigned int OffsetXFenetre, OffsetYFenetre;        // offset en X et Y de la fen�tre
    
} TVipsPULSAR;

// Pr�d�claration des fonctions export�es
int __stdcall PULSAR_Init(TVipsPULSAR *TVP,           // initialisation de la carte PULSAR
                         char *NomFichierDCF,
                         unsigned int NbBuffersDma,
                         unsigned int *XSize, unsigned int *YSize,
                         BOOL UserBitOut,
                         BOOL DoubleBuffer,
                         unsigned int DureeTimeOut,
                         char Resolution,
                         BOOL Fenetre,
                         unsigned int LargeurFenetre, unsigned int HauteurFenetre,
                         unsigned int OffsetXFenetre, unsigned int OffsetYFenetre);
int __stdcall PULSAR_Acquisition(TVipsPULSAR *TVP);   // acquisition d'une image
int __stdcall PULSAR_RecupereImage(TVipsPULSAR *TVP, unsigned char *RawData); // recup�re une image captur�e
int __stdcall PULSAR_Fin(TVipsPULSAR *TVP);           // lib�re la carte PULSAR

/********************************************************/

#define GRAB_START                  11                   // �v�nement "D�but de capture d'image"
#define GRAB_END                    12                   // �v�nement "Fin de capture d'image"                           

#define ACQUISITION_SYNCHRONE       1                   // acquisition synchrone                       
#define ACQUISITION_ASYNCHRONE      2                   // acquisition asynchrone

#define INITIALISATION_COMPLETE     1                   // initialisation compl�te de la carte Pulsar (apr�s un boot)
#define INITIALISATION_PARTIELLE    2                   // initialisation partielle de la carte Pulsar

#define VIDEO_DLL_NON_TROUVEE_ERR           271         // dll non pr�sente ou non accessible
#define VIDEO_FONCTION_NON_TROUVEE_ERR      272         // fonction non trouv�e � l'int�rieur de la dll
#define VIDEO_DLL_NON_LIBERABLE_ERR         273         // erreur pendant la lib�ration dynamique de la dll
#define VIDEO_MDIG_HOOK_FUNCTION_ERR        274         // erreur lors de l'appel de MdigHookFunction                 
#define VIDEO_EVENEMENT_NON_TROUVE_ERR      275         // �v�nement non trouv�

int __stdcall VipsGrab_ChargeFonctionsMagic();  // charge dynamiquement les fonctions de la DLL Magic
int __stdcall VipsGrab_ChargeFonctionsMil();    // charge dynamiquement les fonctions des DLLs Mil
int __stdcall VipsGrab_DechargeDlls();          // d�charge dynamiquement les Dlls

int __stdcall PULSAR_Init2(TVipsPULSAR *TVP,        // nouvelle focntion d'initialisation de la carte PULSAR
                         char *NomFichierDCF,
                         unsigned int NbBuffersDma,
                         unsigned int *XSize, unsigned int *YSize,
                         BOOL UserBitOut,
                         BOOL DoubleBuffer,
                         unsigned int DureeTimeOut,
                         char Resolution,
                         BOOL Fenetre,
                         unsigned int LargeurFenetre, unsigned int HauteurFenetre,
                         unsigned int OffsetXFenetre, unsigned int OffsetYFenetre,
                         char ModeAcquisition,
                         char TypeInitialisation);
int __stdcall PULSAR_DemandeAcquisitionAsynchrone(TVipsPULSAR *TVP);                // demande une acquisition asynchrone � la carte
int __stdcall PULSAR_NotificationDeFinAcquisitionAsynchrone(TVipsPULSAR *TVP);      // notification de fin d'acquisition asynchrone
int __stdcall PULSAR_AttacheFonctionEvenementielle(TVipsPULSAR *TVP, char TypeEvenement, MDIGHOOKFCTPTR PointeurFonction, void MPTYPE *UserDataPtr);   // attache une fonction � un �v�nement

/********************************************************/

#define VIDEO_DPK_INIT_PCK_ERR                  281             // erreur DPK_InitPCK
#define VIDEO_DPK_PCK_SELECT_XPG_ERR            282             // erreur DPK_PCKSelectXPG
#define VIDEO_DPK_INIT_XPG_ERR                  283             // erreur DPK_InitXPG
#define VIDEO_DPF_LOAD_CPF_ERR                  284             // erreur DPF_LoadCPF
#define VIDEO_DBF_SELECT_CPS_ERR                285             // erreur DBF_SelectCPS
#define VIDEO_DBF_GET_GRAB_WINDOW_ERR           286             // erreur DBF_GetGrabWindow
#define VIDEO_DPK_USE_BUS_MASTERING_MODE_ERR    287             // erreur DPK_UseBusMasteringMode 
#define VIDEO_DBF_SET_FRAME_COUNT               288             // erreur DBF_SetFrameCount
#define VIDEO_DBK_XDC_SET_COM_PORT_NUMBER_ERR   289             // erreur DBK_XDCSetComPortNumber
#define VIDEO_DPW_RT_CREATE_BUFFER_ERR          290             // erreur DPW_RTCreateBuffer
#define VIDEO_DPW_RT_REGISTER_BUFFER_ERR        291             // erreur DPW_RTRegisterBuffer
#define VIDEO_DPW_RT_START_ERR                  292             // erreur DPW_RTStart
#define VIDEO_DPW_RT_GET_GRAB_STATUS_ERR        293             // erreur DPW_RTGetGrabStatus
#define VIDEO_DPW_RT_KILL_ERR                   294             // erreur DPW_RTKill
#define VIDEO_DPW_RT_DESTROY_BUFFER_ERR         295             // erreur DPW_RTDestroyBuffer
#define VIDEO_DBF_FREE_CPS_ERR                  296             // erreur DBF_FreeCPS
#define VIDEO_LARGEUR_CAPTURE_ERR               297             // erreur sur la largeur de la capture
#define VIDEO_HAUTEUR_CAPTURE_ERR               298             // erreur sur la hauteur de la capture

#define VIDEO_DBF_SET_GRAB_WINDOW_ERR           299             // erreur DBF_SetGrabWindow

// d�finition de la structure TVipsLPG
// structure permettant la gestion des cartes DipixLPG
typedef struct
{
    long ErreurDIPIX;                                           // num�ro d'erreur DIPIX

    char NomFichierCPF[TAILLE_MAX_CHAINE];                      // nom du fichier CPF
                                                                // contenant les param�tres de
                                                                // r�glages de la cam�ra
    long NbBuffersDma;      // nombre de buffers DMA utilis�s                                

    char Resolution;        // r�solution de la capture

    // fen�tre d'acquisition initiale
    long StartPixel, NombrePixels;          
    long StartLigne, NombreLignes;

    BOOL Fenetre;               // pour utiliser le fen�trage ou non
    unsigned int LargeurFenetre, HauteurFenetre;        // largeur et hauteur de la fen�tre
    unsigned int OffsetXFenetre, OffsetYFenetre;        // offset en X et Y de la fen�tre

    long TailleImage;           // taille de l'image en octet
                        
    long HandleCPS;             // handle de la structure CPS (structure interne DIPIX)        
    void *PointeurBufferDma;    // pointeur sur le buffer Dma
    long TailleBufferDma;       // taille en octet du buffer Dma

    long CompteurImage;             // nombre d'images captur�es depuis le dernier d�marrage d'une acquisition
    long CompteurImagePrecedent;    // pr�c�dent nombre d'images captur�es depuis le dernier d�marrage d'une acquisition
    int IndexImageDansBufferDma;    // index de l'image captur�e dans le buffer Dma

} TVipsLPG;

// initialisation de la carte DipixLPG
int __stdcall LPG_Init(TVipsLPG *TVLPG, char *NomFichierCPF, long NbBuffersDma,
                            long *XSize,
                            long *YSize,
                            char Resolution,
                            BOOL Fenetre,
                            unsigned int LargeurFenetre, unsigned int HauteurFenetre,
                            unsigned int OffsetXFenetre, unsigned int OffsetYFenetre);
// d�marre la capture
int __stdcall LPG_DemarreCapture(TVipsLPG *TVLPG);
// attend la fin d'une capture
int __stdcall LPG_AttenteFinImageCapturee(TVipsLPG *TVLPG, long *NbImagesCapturees, int *IndexPremiereImageDansBufferDma);
// recup�re une image captur�e
int __stdcall LPG_RecupereImage(TVipsLPG *TVLPG, unsigned char *RawData, int IndexImageDansBufferDma);
// arr�te la capture
int __stdcall LPG_FinCapture(TVipsLPG *TVLPG);
// lib�re la carte DipixLPG
int __stdcall LPG_Fin(TVipsLPG *TVLPG);    
// charge dynamiquement les fonctions de la dll DipixLPG.
int __stdcall VipsGrab_ChargeFonctionsDipixLPG();

#ifdef __cplusplus
}
#endif

#endif /* __VIPSGRAB_H__ */