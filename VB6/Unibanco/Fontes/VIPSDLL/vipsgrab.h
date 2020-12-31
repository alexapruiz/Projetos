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
/*          - démarrage de la capture                             */
/*          - acquisition d'une image                             */
/*          - recupération de l'image capturée                    */
/*          - arrêt de la capture                                 */
/*          - choix du mode d'affichage vidéo (MAGIC seulement)   */
/* Version : v1.20                                                */
/* Développée par : M. Tola                                       */
/*----------------------------------------------------------------*/
/* MODIFICATIONS                                                  */
/*----------------------------------------------------------------*/
/* Date de la modification : 11/12/97                             */
/* Code : - LEX                                                   */
/*        - gestion de la carte PULSAR                            */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 22/12/97                             */
/* Code : - LEX                                                   */
/*        - modularité de la DLL au niveau des cartes utilisées   */
/*        - renommage de la DLL en VipsGrab.dll                   */
/*        - renomage des fonctions GetDLLVideoVersion en          */
/*          GetDLLVipsGrabVersion et Video_Erreur en              */
/*          VipsGrab_Erreur                                       */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 16/01/98                             */
/* Code : - LEX                                                   */
/*        - gestion 100 et 200 DPI                                */
/*        - gestion fenêtrage de la capture                       */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 20/03/98                             */
/* Code : - LEX                                                   */
/*        - chargement dynamique des dlls                         */
/*        - acquisition asynchrone                                */
/*        - fonctions évènementielles                             */
/*        - création de la fct. Pulsar_Init2                      */
/*        - Pulsar_Init appelle désormais Carte_Pulsar_Init       */
/*        - MAGIC_DemarreCapture et Magic_FinCapture retournent   */
/*          désormais un int                                      */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 21/04/98                             */
/* Code : - LEX                                                   */
/*        - VipsGrab_DechargeDlls modifiée pour mieux gérer le    */
/*          déchargement dynamique des Dlls                       */
/*        - adaptation de PULSAR_AttacheFonctionEvenementielle    */
/*          à Mil-Lite 5.1 (gestion des évènements GRAB_START     */
/*          et GRAB_END)                                          */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 29/04/98                             */
/* Code : - LEX                                                   */
/*        - gestion des allocations dynamiques des buffers DMA    */
/*          sur la carte PULSAR ou en zone non swappable          */
/*          (cf. Carte_Pulsar_Init)                               */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 26/08/98                             */
/* Code : - LDB                                                   */
/*        - correction positionnement fenêtrage acquisition en    */
/*          100 DPI                                               */
/*                                                                */
/* Modifié par : M. Tola                                          */
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

// nombre maximal de buffers DMA autorisés
#define NB_MAX_BUFFERS_DMA            10

// taille maximale d'une chaine de caractères
#define TAILLE_MAX_CHAINE             256

// "profondeur" d'un plan image,
// en 256 niveaux de gris, 1 pixel = 1 octet 
#define DATA_DEPTH                    8

// #define des erreurs de la DLL
#define VIDEO_BUFFER_DMA_PLEIN_ERR                221   // buffer DMA plein
#define VIDEO_TROP_DE_BUFFERS_DMA_DEMANDES_ERR    222   // trop de buffers DMA demandés par rapport au nombre autorisé
#define VIDEO_BUFFER_DMA_VIDE_ERR                 223   // buffer DMA vide

// Prédéclaration des fonctions exportées
void __stdcall GetDLLVipsGrabVersion(char *Version);   // retourne le numéro de la version de la DLL sous forme d'une chaine de caractères
int __stdcall VipsGrab_Erreur(unsigned int NumeroErreur, char *ChaineErreur);  // retourne la chaine de caractères correspondante au numéro d'erreur passé en paramètre

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

// #define des différents modes d'affichages possibles
#define VGA                 11      // réalise uniquement l'affichage VGA
#define IMAGE               12      // réalise uniquement l'affichage du buffer image de la carte d'acquisition
#define VGA_ET_IMAGE        13      // superpose l'affichage VGA et le buffer image quand la couleur de fond est 0 donc noire

// #define des erreurs de la DLL
#define VIDEO_INIT_DEFAULT_ERR                201   // erreur retournée par mlInitDefault
#define VIDEO_CAM_LOAD_ERR                    202   // erreur retournée par mlCamLoad
#define VIDEO_DMA_BUFFER_ALLOC_ERR            203   // erreur retournée par mlDmaBufferAlloc
#define VIDEO_DMA_BUFFER_GET_PTR_ERR          204   // erreur retournée par mlDmaBufferGetPtr
#define VIDEO_DMA_BUFFER_SELECT_ERR           205   // erreur retournée par mlDmaBufferSelect
#define VIDEO_DMA_BUFFER_DESELECT_ERR         206   // erreur retournée par mlDmaBufferDeselect
#define VIDEO_DMA_BUFFER_FREE_ERR             207   // erreur retournée par mlDmaBufferFree
#define VIDEO_DMA_WAIT_ERR                    208   // erreur retournée par mlDmaWait
#define VIDEO_DMA_READ_TRANS_ERR              209   // erreur retournée par mlDmaReadTrans
#define VIDEO_DMA_BUFFER_READ_AREA_ERR        210   // erreur retournée par mlDmaBufferReadArea

#define VIDEO_CAM_USER_BIT_OUT_STATE_ERR      211   // erreur retournée par mlCamUserBitOutState
#define VIDEO_CAM_USER_BIT_OUT_ERR            212   // erreur retournée par mlCamUserBitOut

// cf. #define 221, 222, 223 ci-dessus
#define VIDEO_MODE_AFFICHAGE_INCONNU              224   // mode d'affichage inconnu

// définition de la structure TVipsMAGIC
// structure permettant la gestion des cartes MAGIC
typedef struct
{
    TCAMERA    Camera;      // structure contenant les paramètres de
                            // réglages de la caméra

    char NomFichierDCF[TAILLE_MAX_CHAINE];      // nom du fichier DCF
                                                // contenant les paramètres de
                                                // réglages de la caméra

    BOOL UserBitOut;        // booléen permettant de générer un front montant
                            // (sur le UserBit 0 de la carte MAGIC) après chaque acquisition dans MAGIC_Acquisition

    BOOL DoubleBuffer;      // booléen permettant de gérer le "double buffering"

    unsigned int LargeurImage, HauteurImage;    // largeur et hauteur d'une image capturée (en pixels)
    
    unsigned int NbBuffersDma;          // nombre de buffers DMA utilisés

    short IdentificateurBufferDma[NB_MAX_BUFFERS_DMA];      // tableau d'identificateurs de chaque buffer DMA,
                                                            // cet identificateur est utile pour les fonctions de transfert DMA                                                                    

    unsigned char *AdresseBufferDma[NB_MAX_BUFFERS_DMA];    // tableau de pointeurs de chaque buffer DMA

    unsigned int NbElemsDansBufferDma;      // nombre d'images dans le buffer DMA
    unsigned int PointeurInBufferDma;       // index d'entrée dans le buffer DMA
    unsigned int PointeurOutBufferDma;      // index de sortie dans le buffer DMA

    // numéro de page virtuelle (gestion du "double buffering")
    short Page;

    unsigned char        *future[5];  // tableau de pointeurs pour "réserver de la place"
                                      // pour des extensions futures et amener la taille
                                      // de la structure à 832 octets
    
} TVipsMAGIC;

// Prédéclaration des fonctions exportées
int __stdcall MAGIC_Init(TVipsMAGIC *TVM,           // initialisation de la carte MAGIC
                         char *NomFichierDCF,
                         unsigned int NbBuffersDma,
                         unsigned int *XSize, unsigned int *YSize,
                         BOOL UserBitOut,
                         BOOL DoubleBuffer);
int __stdcall MAGIC_DemarreCapture(void);           // démarre la capture
int __stdcall MAGIC_Acquisition(TVipsMAGIC *TVM);   // acquisition d'une image
int __stdcall MAGIC_RecupereImage(TVipsMAGIC *TVM, unsigned char *RawData); // recupère une image capturée
int __stdcall MAGIC_FinCapture(void);               // arrête la capture
int __stdcall MAGIC_Fin(TVipsMAGIC *TVM);           // libère la carte MAGIC
int __stdcall MAGIC_Affichage(int ModeAffichage);   // change le mode d'affichage

/********************************************************/

// include de mil.dll et milpul.dll : gestion des cartes PULSAR
#include <mil.h>

#define RES100DPI      1            // capture en 100 dpi
#define RES200DPI      2            // capture en 200 dpi : nécessite un fichier DCF adéquat

#define VIDEO_MAPP_ALLOC_ERR        231     // erreur retournée par MappAlloc
#define VIDEO_MSYS_ALLOC_ERR        232     // erreur retournée par MsysAlloc
#define VIDEO_MDIG_ALLOC_ERR        233     // erreur retournée par MdigAlloc
#define VIDEO_MDIG_CONTROL_ERR      234     // erreur retournée par MdigControl
#define VIDEO_MDIG_CHANNEL_ERR      235     // erreur retournée par MdigChannel
#define VIDEO_BUF_ALLOC_2D_ERR      236     // erreur retournée par MbufAlloc2d
#define VIDEO_MBUF_FREE_ERR         237     // erreur retournée par MbufFree
#define VIDEO_MDIG_FREE_ERR         238     // erreur retournée par MdigFree
#define VIDEO_MSYS_FREE_ERR         239     // erreur retournée par MsysFree
#define VIDEO_MAPP_FREE_ERR         240     // erreur retournée par MappFree
#define VIDEO_MDIG_HALT_ERR         241     // erreur retournée par MdigHalt
#define VIDEO_MDIG_GRAB_ERR         242     // erreur retournée par MdigGrab
#define VIDEO_MDIG_GRAB_WAIT_ERR    243     // erreur retournée par MdigGrabWait
#define VIDEO_MBUF_GET_2D_ERR       244     // erreur retournée par MbufGet2d
#define VIDEO_MAPP_CONTROL_ERR      245     // erreur retournée par MappControl
#define VIDEO_MDIG_INQUIRE_ERR      246     // erreur retournée par MdigInquire

#define VIDEO_MDIG_GRAB_TIMEOUT_ERR 251     // erreur de TimeOut pendant la capture

#define VIDEO_LARGEUR_FENETRE_ERR    261    // erreur sur la largeur de la fenêtre
#define VIDEO_HAUTEUR_FENETRE_ERR    262    // erreur sur la hauteur de la fenêtre
#define VIDEO_LARGEUR_IMAGE_ERR      263    // erreur sur la largeur de l'image
#define VIDEO_HAUTEUR_IMAGE_ERR      264    // erreur sur la hauteur de l'image
#define VIDEO_OFFSETX_FENETRE_ERR    265    // erreur sur l'offset en X de la fenêtre
#define VIDEO_OFFSETY_FENETRE_ERR    266    // erreur sur l'offset en Y de la fenêtre
#define VIDEO_FENETRE_ERR            267    // erreur, mode de fenêtrage non actif

// définition de la structure TVipsPULSAR
// structure permettant la gestion des cartes PULSAR
typedef struct
{
    char NomFichierDCF[TAILLE_MAX_CHAINE];      // nom du fichier DCF
                                                // contenant les paramètres de
                                                // réglages de la caméra

    // non utilisé pour le moment
    BOOL UserBitOut;        // booléen permettant de générer un front montant
                            // (sur le UserBit 0 de la carte PULSAR) après chaque acquisition dans PULSAR_Acquisition

    // non utilisé pour le moment
    BOOL DoubleBuffer;      // booléen permettant de gérer le "double buffering"

    unsigned int LargeurImage, HauteurImage;    // largeur et hauteur d'une image capturée (en pixels)
    
    unsigned int NbBuffersDma;          // nombre de buffers DMA utilisés

    unsigned int NbElemsDansBufferDma;      // nombre d'images dans le buffer DMA
    unsigned int PointeurInBufferDma;       // index d'entrée dans le buffer DMA
    unsigned int PointeurOutBufferDma;      // index de sortie dans le buffer DMA

    // non utilisé pour le moment
    short Page;     // numéro de page virtuelle (gestion du "double buffering")

    unsigned char        *future[3];  // tableau de pointeurs pour "réserver de la place"
                                      // pour des extensions futures et amener la taille
                                      // de la structure à 384 octets

    /****************/

    unsigned int DureeTimeOut;      // TimeOut de la capture d'image

    MIL_ID MilApplication;   // identificateur de "l'application" MIL
    MIL_ID MilSysteme;       // identificateur du "système" MIL
    MIL_ID MilCamera;        // identificateur de la "caméra" MIL

    MIL_ID IdentificateurBufferDma[NB_MAX_BUFFERS_DMA];     // identificateurs des "buffers DMA" MIL

    char Resolution;        // résolution de la capture

    BOOL Fenetre;            // pour utiliser le fenêtrage ou non
    unsigned int LargeurFenetre, HauteurFenetre;        // largeur et hauteur de la fenêtre
    unsigned int OffsetXFenetre, OffsetYFenetre;        // offset en X et Y de la fenêtre
    
} TVipsPULSAR;

// Prédéclaration des fonctions exportées
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
int __stdcall PULSAR_RecupereImage(TVipsPULSAR *TVP, unsigned char *RawData); // recupère une image capturée
int __stdcall PULSAR_Fin(TVipsPULSAR *TVP);           // libère la carte PULSAR

/********************************************************/

#define GRAB_START                  11                   // évènement "Début de capture d'image"
#define GRAB_END                    12                   // évènement "Fin de capture d'image"                           

#define ACQUISITION_SYNCHRONE       1                   // acquisition synchrone                       
#define ACQUISITION_ASYNCHRONE      2                   // acquisition asynchrone

#define INITIALISATION_COMPLETE     1                   // initialisation complète de la carte Pulsar (après un boot)
#define INITIALISATION_PARTIELLE    2                   // initialisation partielle de la carte Pulsar

#define VIDEO_DLL_NON_TROUVEE_ERR           271         // dll non présente ou non accessible
#define VIDEO_FONCTION_NON_TROUVEE_ERR      272         // fonction non trouvée à l'intérieur de la dll
#define VIDEO_DLL_NON_LIBERABLE_ERR         273         // erreur pendant la libération dynamique de la dll
#define VIDEO_MDIG_HOOK_FUNCTION_ERR        274         // erreur lors de l'appel de MdigHookFunction                 
#define VIDEO_EVENEMENT_NON_TROUVE_ERR      275         // évènement non trouvé

int __stdcall VipsGrab_ChargeFonctionsMagic();  // charge dynamiquement les fonctions de la DLL Magic
int __stdcall VipsGrab_ChargeFonctionsMil();    // charge dynamiquement les fonctions des DLLs Mil
int __stdcall VipsGrab_DechargeDlls();          // décharge dynamiquement les Dlls

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
int __stdcall PULSAR_DemandeAcquisitionAsynchrone(TVipsPULSAR *TVP);                // demande une acquisition asynchrone à la carte
int __stdcall PULSAR_NotificationDeFinAcquisitionAsynchrone(TVipsPULSAR *TVP);      // notification de fin d'acquisition asynchrone
int __stdcall PULSAR_AttacheFonctionEvenementielle(TVipsPULSAR *TVP, char TypeEvenement, MDIGHOOKFCTPTR PointeurFonction, void MPTYPE *UserDataPtr);   // attache une fonction à un évènement

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

// définition de la structure TVipsLPG
// structure permettant la gestion des cartes DipixLPG
typedef struct
{
    long ErreurDIPIX;                                           // numéro d'erreur DIPIX

    char NomFichierCPF[TAILLE_MAX_CHAINE];                      // nom du fichier CPF
                                                                // contenant les paramètres de
                                                                // réglages de la caméra
    long NbBuffersDma;      // nombre de buffers DMA utilisés                                

    char Resolution;        // résolution de la capture

    // fenêtre d'acquisition initiale
    long StartPixel, NombrePixels;          
    long StartLigne, NombreLignes;

    BOOL Fenetre;               // pour utiliser le fenêtrage ou non
    unsigned int LargeurFenetre, HauteurFenetre;        // largeur et hauteur de la fenêtre
    unsigned int OffsetXFenetre, OffsetYFenetre;        // offset en X et Y de la fenêtre

    long TailleImage;           // taille de l'image en octet
                        
    long HandleCPS;             // handle de la structure CPS (structure interne DIPIX)        
    void *PointeurBufferDma;    // pointeur sur le buffer Dma
    long TailleBufferDma;       // taille en octet du buffer Dma

    long CompteurImage;             // nombre d'images capturées depuis le dernier démarrage d'une acquisition
    long CompteurImagePrecedent;    // précédent nombre d'images capturées depuis le dernier démarrage d'une acquisition
    int IndexImageDansBufferDma;    // index de l'image capturée dans le buffer Dma

} TVipsLPG;

// initialisation de la carte DipixLPG
int __stdcall LPG_Init(TVipsLPG *TVLPG, char *NomFichierCPF, long NbBuffersDma,
                            long *XSize,
                            long *YSize,
                            char Resolution,
                            BOOL Fenetre,
                            unsigned int LargeurFenetre, unsigned int HauteurFenetre,
                            unsigned int OffsetXFenetre, unsigned int OffsetYFenetre);
// démarre la capture
int __stdcall LPG_DemarreCapture(TVipsLPG *TVLPG);
// attend la fin d'une capture
int __stdcall LPG_AttenteFinImageCapturee(TVipsLPG *TVLPG, long *NbImagesCapturees, int *IndexPremiereImageDansBufferDma);
// recupère une image capturée
int __stdcall LPG_RecupereImage(TVipsLPG *TVLPG, unsigned char *RawData, int IndexImageDansBufferDma);
// arrête la capture
int __stdcall LPG_FinCapture(TVipsLPG *TVLPG);
// libère la carte DipixLPG
int __stdcall LPG_Fin(TVipsLPG *TVLPG);    
// charge dynamiquement les fonctions de la dll DipixLPG.
int __stdcall VipsGrab_ChargeFonctionsDipixLPG();

#ifdef __cplusplus
}
#endif

#endif /* __VIPSGRAB_H__ */