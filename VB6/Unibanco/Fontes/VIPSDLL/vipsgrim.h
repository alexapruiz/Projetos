/******************************************************************/
/*                             LIBRAIRIE                          */
/*----------------------------------------------------------------*/
/* Librairie : vipsgrim.dll                                       */
/* Titre : - gestion des cartes d'acquisitions MAGIC et           */
/*           PULSAR                                               */
/*         - gestion images BMP, ECF, TIFF G4, JPEG               */
/*         - captures d'images avec un LC13                       */
/* Contenu: - initialisation des cartes                           */
/*          - démarrage de la capture                             */
/*          - acquisition d'une image                             */
/*          - recupération de l'image capturée                    */
/*          - arrêt de la capture                                 */
/*          - choix du mode d'affichage vidéo (MAGIC seulement)   */
/*                                                                */
/*                                                                */
/*          - création d'une image BMP                            */
/*          - rotation (90°, 180° et 270°) d'une image BMP        */
/*          - symétrie horizontale et verticale d'une image BMP   */
/*          - compression des images BMP en ECF                   */
/*          - décompression des images ECF en BMP                 */
/*          - compression des images BMP en ECF et écriture       */
/*            directe                                             */
/*          - découpe d'une image dans une autre et création d'une*/
/*            nouvelle image                                      */
/*          - auto-suppression des bords noirs de l'image capturée*/
/*          - capture d'un étalon lumineux sur une image          */
/*          - lecture et écriture d'un étalon lumineux sur le     */
/*            disque                                              */
/*          - correction lumineuse d'une image (à partir de       */
/*            l'étalon et/ou verticale)                           */
/*          - compression des images BMP en JPG et écriture       */
/*            directe                                             */
/*          - décompression des images JPG en BMP                 */
/*          - compression des images BMP en TIFF G4 et écriture   */
/*            directe                                             */
/*          - décompression des images TIFF G4 en BMP             */
/* Version : v1.20                                                */
/* Développée par : M. Tola                                       */
/*----------------------------------------------------------------*/
/* MODIFICATIONS                                                  */
/*----------------------------------------------------------------*/
/* Date de la modification : 21/04/98                             */
/* Code : - LEX                                                   */
/*        - la priorité de ...ReglePrioriteProgrammeEtMain        */
/*          est désormais NORMALE                                 */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 30/04/98                             */
/* Code : - LEX                                                   */
/*        - gestion de la correction lumineuse et de l'étalon     */
/*          dans VipsGrimLC13_Init et                             */
/*          VipsGrimLC13_TacheCompressionEtEcriture               */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 26/05/98                             */
/* Code : - LEX                                                   */
/*        - gestion deuxième fonction UserData                    */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 16/07/98                             */
/* Code : - LEX                                                   */
/*        - gestion fonction UserData en __stdcall pour VB        */
/*        - rajout VipsGrimLC13_AttenteNImagesPresentesDansBuffer */
/*        - rajout VipsGrimLC13_AttenteNFichiersPresentsDansBuffer*/
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 28/08/98                             */
/* Code : - LEX                                                   */
/*        - gestion des noms de fichiers avec suffixe déjà        */
/*          présents                                              */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 16/09/98                             */
/* Code : - LEX                                                   */
/*        - gestion carte DipixLPG                                */
/*                                                                */
/* Modifié par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 17/09/98                             */
/* Code : - LEX                                                   */
/*        - gestion du seuillage automatique                      */
/*                                                                */
/* Modifié par : M. Tola                                          */
/******************************************************************/

#ifndef __VIPSGRIM_H__
#define __VIPSGRIM_H__

#ifdef __cplusplus
extern "C" {
#endif

// include de Windows
#include <stdio.h>
#include <conio.h>
#include <stdlib.h>
#include <string.h>
#include <windows.h>

// gestion des buffers images et des buffers fichiers
#include "buff_def.h"

// constantes prédéfinies
#define SANS_ECRITURE                           0
#define FICHIER_BMP                             1
#define FICHIER_ECF                             2
#define FICHIER_JPG                             3
#define FICHIER_TIFF                            4
#define FICHIER_BMP_NB                          5
#define FICHIER_JPG_LEADTOOLS                   6
#define FICHIER_JPG_IJG                         7

// constantes prédéfinies
#define RECTO                                   1
#define VERSO                                   2
#define FACE_NULLE                              3

// taille maximale d'un nom de fichier
#define TAILLE_MAX_FICHIER                      256

// taille maximale d'un chemin
#define TAILLE_MAX_CHEMIN                       1024

// pour savoir si la fonction UserData est appelée avant l'écriture de
// l'image sur le disque, ou après
#define USERDATA_AVANT_ECRITURE                 11
#define USERDATA_APRES_ECRITURE                 12

// #define des erreurs
#define VIPSGRIMLC13_CREATE_SEMAPHORE_ERR                               601     // on ne peut pas créer le sémaphore
#define VIPSGRIMLC13_RELEASE_SEMAPHORE_ERR                              602     // on ne peut pas "incrémenter" le sémaphore
#define VIPSGRIMLC13_WAIT_FOR_SINGLE_OBJECT_ERR                         603     // erreur pendant un "wait" sur un sémaphore
#define VIPSGRIMLC13_BEGIN_THREAD_ERR                                   604     // erreur pendant un lancement de thread
#define VIPSGRIMLC13_CLOSE_HANDLE_ERR                                   605     // erreur pendant la libération d'un sémaphore
#define VIPSGRIMLC13_TIMEOUT_SEMAPHORE_BUFFERS_DMA_LIBRES               606     // timeout du wait du sémaphore SemaphoreNbBuffersDmaLibres
#define VIPSGRIMLC13_TIMEOUT_SEMAPHORE_BUFFERS_IMAGES_LIBRES            607     // timeout du wait du sémaphore SemaphoreNbBuffersImagesLibres
#define VIPSGRIMLC13_ATTACHE_FONCTION_EVENEMENTIELLE_ERR                608     // erreur lors de l'attachement d'une fonction à un évènement
#define VIPSGRIMLC13_BUFFER_IMAGES_PRENDRE_ERR                          609     // erreur lors de la récuperation d'une image à partir du buffer image
#define VIPSGRIMLC13_BUFFER_FICHIERS_PRENDRE_ERR                        610     // erreur lors de la récuperation d'un fichier à partir du buffer fichier
#define VIPSGRIMLC13_PTR_FONCTION_USER_DATA_BUFFER_FICHIERS_ERR         611     // erreur au sein de la fonction UserDataBufferFichiers
#define VIPSGRIMLC13_DLL_VIPSPROD_NON_TROUVEE_ERR                           612 // dll VipsProd non trouvée
#define VIPSGRIMLC13_FONCTION_AUTORISATION_ACCES_PRODUIT_NON_TROUVEE_ERR    613 // fonction AutorisationAccesProduit non trouvée dans la DLL VipsProd
#define VIPSGRIMLC13_FREE_LIBRARY_ERR                                       614 // erreur de libération de la DLL VipsProd
#define VIPSGRIMLC13_AUTORISATION_ACCES_DLL_REFUSEE_ERR                     615 // autorisation d'accès refusée

#define VIPSGRIMLC13_RELEASE_SEMAPHORE_NB_BUFFERS_DMA_LIBRES_ERR                616 // erreur release semaphore nb buffers dma libres
#define VIPSGRIMLC13_RELEASE_SEMAPHORE_NB_BUFFERS_IMAGES_LIBRES_ERR             617 // erreur release semaphore nb buffers images libres
#define VIPSGRIMLC13_RELEASE_SEMAPHORE_NB_IMAGES_PRESENTES_DANS_BUFFER_ERR      618 // erreur release semaphore nb images presentes dans buffer
#define VIPSGRIMLC13_RELEASE_SEMAPHORE_NB_FICHIERS_PRESENTS_DANS_BUFFER_ERR     619 // erreur release semaphore nb fichiers presents dans buffer
#define VIPSGRIMLC13_TIMEOUT_SEMAPHORE_IMAGES_PRESENTES_DANS_BUFFER             620 // timeout du wait du sémaphore SemaphoreNbImagesPresentesDansBuffer
#define VIPSGRIMLC13_TIMEOUT_SEMAPHORE_FICHIERS_PRESENTS_DANS_BUFFER            621 // timeout du wait du sémaphore SemaphoreNbFichiersPresentsDansBuffer

// déclaration d'un nouveau type : pointeur sur FonctionUserDataBufferFichiers
typedef int (* PTR_FonctionUserDataBufferFichiers)(char MomentAppel, unsigned int InformationFace,
                                                    char *NomFichier, TVipsBMP *TVB, BOOL *Ecriture,
                                                    USERDATA_BUFFER_FICHIERS *PUserDataBufferFichiers);

// déclaration d'un nouveau type : pointeur sur FonctionUserDataBufferFichiers_STDCALL (pour VB)
typedef int (__stdcall * PTR_FonctionUserDataBufferFichiers_STDCALL)(char MomentAppel, unsigned int InformationFace,
                                                                        char *NomFichier, TVipsBMP *TVB, BOOL *Ecriture,
                                                                        USERDATA_BUFFER_FICHIERS *PUserDataBufferFichiers);

// déclaration de la structure TVipsStatut
// ne pas accéder à ces champs directement
typedef struct
{
    unsigned int NbErreursSurvenues;                        // nb d'erreurs survenues
    
    unsigned int NumeroErreur;                              // numéro de l'erreur
    char NomFonctionRetournantErreur[TAILLE_MAX_CHAINE];    // fonction retournant l'erreur
    char NomFonction[TAILLE_MAX_CHAINE];                    // fonction contenant la fonction retournant l'erreur

    unsigned int ErreurGetLast;                             // statut de ReleaseSemaphore SemaphoreNbBuffersDmaLibres
} TVipsStatut;

// déclaration de la structure TVipsLC13
// ne pas accéder à ces champs directement
typedef struct
{
    BOOL ArretTaches;               // booléen permettant de gérer l'arrêt des tâches

    unsigned int NbBuffersDma;      // nb de buffers DMA utilisés par la carte vidéo

    char NomFichierDCF[TAILLE_MAX_FICHIER];     // nom du fichier DCF, contient les paramètres de la caméra

    TVipsPULSAR TVP;                // structure du type TVipsPULSAR, gestion de la carte vidéo PULSAR

    unsigned int NbBuffersImages;   // nb de buffers images utilisés
    TVipsBufferImages TVBI;         // structure du type TVipsBufferImages, gestion des buffers tournants

    char CheminStockageImages[TAILLE_MAX_CHEMIN];   // chemin de stockage des images
    char ResolutionImages;                          // RES100DPI ou RES200DPI
    char TypeFichier;                       // type de fichier (SANS_ECRITURE, FICHIER_BMP, FICHIER_ECF, FICHIER_JPG, FICHIER_TIFF, FICHIER_BMP_NB)
    unsigned int QualiteCompressionECF;     // qualité de compression pour les fichiers ECF
    unsigned int QualiteCompressionJPG;     // qualité de compression pour les fichiers JPG
    unsigned char SeuilBinarisationTIFF;    // seuil de binarisation des images TIFF

    // 2 opérations possibles pour chaque face de l'image avant la compression
    // et la sauvegarde
    char OrientationRecto1, OrientationVerso1;
    char OrientationRecto2, OrientationVerso2;

    // position et taille de l'acquisition fenêtrée (si utilisée)
    BOOL AcquisitionFenetree;
    unsigned int LargeurFenetre, HauteurFenetre;
    unsigned int OffsetXFenetre, OffsetYFenetre;

    // handle des sémaphores
    HANDLE SemaphoreNbBuffersDmaLibres;
    HANDLE SemaphoreNbBuffersImagesLibres;
    HANDLE SemaphoreNbImagesPresentesDansBuffer;
    HANDLE SemaphoreNbFichiersPresentsDansBuffer;

    TVipsStatut TVS;        // statut d'erreur

    // nb de grabs en attente
    unsigned int NbGrabsEnAttente;

    unsigned int NbBuffersFichiers;     // nb de buffers fichiers utilisés
    TVipsBufferFichiers TVBF;           // structure du type TVipsBufferFichiers, gestion des buffers tournants

    PTR_FonctionUserDataBufferFichiers Ptr_FonctionUserDataBufferFichiers;                      // pointeur vers la fonction UserDataBufferFichiers
    PTR_FonctionUserDataBufferFichiers_STDCALL Ptr_FonctionUserDataBufferFichiers_STDCALL;      // pointeur vers la fonction UserDataBufferFichiers (pour VB)

    unsigned int TailleZoneUserDataBufferFichiers;      // taille de la zone UserDataBufferFichiers

    BOOL CorrectionLumineuse;                       // correction lumineuse ou pas ?
    char TypeCorrectionLumineuse;                   // CORRECTION_A_PARTIR_ETALON, CORRECTION_VERTICALE ou TOUTES_CORRECTIONS_POSSIBLES
    TVipsEtalon TVE;                                // pointeur vers une structure étalon
    char NomFichierEtalon[TAILLE_MAX_FICHIER];      // nom du fichier étalon

    TVipsLPG TVLPG;                                 // structure du type TVipsLPG, gestion de la carte vidéo DipixLPG
    HANDLE HandleThreadAcquisitionDipixLPG;         // handle de la tâche d'acquisition de la carte DipixLPG
    
} TVipsLC13;

void __stdcall GetDLLVipsGrimVersion(char *Version);    // retourne le numéro de la version de la DLL sous forme d'une chaine de caractères
int __stdcall VipsGrim_Erreur(unsigned int NumeroErreur, char *ChaineErreur);   // retourne la chaine de caractères correspondante au numéro d'erreur passé en paramètre
                                                                                                        // Cette fonction utilise le fichier 'vipsgrim.err'.
void __stdcall VipsGrimLC13_ReglePrioriteProgrammeEtMain();                                             // règle la priorité du programme et de la 'tâche' main
int __stdcall VipsGrimLC13_GetStatus(TVipsLC13 *TVLC13, char *NomFonctionRetournantErreur, char *NomFonction, unsigned int *NumeroErreur);      // retourne un status d'erreur si une erreur s'est produite
int __stdcall VipsGrimLC13_AttenteNBuffersDmaLibres(TVipsLC13 *TVLC13, unsigned int nombre);           // attente de N buffers Dma libres
int __stdcall VipsGrimLC13_AttenteNBuffersImagesLibres(TVipsLC13 *TVLC13, unsigned int nombre);        // attente de N buffers images libres
int __stdcall VipsGrimLC13_SignaleNBuffersDmaLibres(TVipsLC13 *TVLC13, unsigned int nombre);           // signale N buffers Dma libres
int __stdcall VipsGrimLC13_SignaleNBuffersImagesLibres(TVipsLC13 *TVLC13, unsigned int nombre);        // signale N buffers images libres
int __stdcall VipsGrimLC13_SignaleNImagesPresentesDansBuffer(TVipsLC13 *TVLC13, unsigned int nombre);  // signale N images présentes dans le buffer image
int __stdcall VipsGrimLC13_SignaleNFichiersPresentsDansBuffer(TVipsLC13 *TVLC13, unsigned int nombre); // signale N fichiers présents dans le buffer fichiers
int __stdcall VipsGrimLC13_OrdreFichier(TVipsLC13 *TVLC13, unsigned int InformationFace, char *NomFichier, USERDATA_BUFFER_FICHIERS *PUserDataBufferFichiers);    // envoi un ordre fichier à la tâche de compression et d'écriture
int __stdcall VipsGrimLC13_Init(TVipsLC13 *TVLC13,           // initialisation de la structure TVipsLC13
    unsigned int NbBuffersDmaIn,
    unsigned int NbBuffersImagesIn,
    unsigned int NbBuffersFichiersIn,
    char *NomFichierDCFIn,
    char *CheminStockageImagesIn,
    char ResolutionImagesIn,
    char TypeFichierIn,
    unsigned int QualiteCompressionECFIn,
    unsigned int QualiteCompressionJPGIn,
    unsigned char SeuilBinarisationTIFFIn,
    char OrientationRectoIn1, char OrientationRectoIn2,
    char OrientationVersoIn1, char OrientationVersoIn2,
    BOOL AcquisitionFenetreeIn,
    unsigned int LargeurFenetreIn, unsigned int HauteurFenetreIn,
    unsigned int OffsetXFenetreIn, unsigned int OffsetYFenetreIn,
    PTR_FonctionUserDataBufferFichiers Ptr_FonctionUserDataBufferFichiersIn,
    unsigned int TailleZoneUserDataBufferFichiersIn,
    BOOL CorrectionLumineuseIn,
    char TypeCorrectionLumineuseIn,
    char *NomFichierEtalonIn);
int __stdcall VipsGrimLC13_Fin(TVipsLC13 *TVLC13);      // libération de la structure TVipsLC13
int __stdcall VipsGrimLC13_ReinitialiseCartePulsar(TVipsLC13 *TVLC13);      // réinitialise la carte Pulsar

int __stdcall VipsGrimLC13_Init_STDCALL(TVipsLC13 *TVLC13,    // initialisation de la structure TVipsLC13 (pour VB)
    unsigned int NbBuffersDmaIn,
    unsigned int NbBuffersImagesIn,
    unsigned int NbBuffersFichiersIn,
    char *NomFichierDCFIn,
    char *CheminStockageImagesIn,
    char ResolutionImagesIn,
    char TypeFichierIn,
    unsigned int QualiteCompressionECFIn,
    unsigned int QualiteCompressionJPGIn,
    unsigned char SeuilBinarisationTIFFIn,
    char OrientationRectoIn1, char OrientationRectoIn2,
    char OrientationVersoIn1, char OrientationVersoIn2,
    BOOL AcquisitionFenetreeIn,
    unsigned int LargeurFenetreIn, unsigned int HauteurFenetreIn,
    unsigned int OffsetXFenetreIn, unsigned int OffsetYFenetreIn,
    PTR_FonctionUserDataBufferFichiers_STDCALL Ptr_FonctionUserDataBufferFichiersIn_STDCALL,
    unsigned int TailleZoneUserDataBufferFichiersIn,
    BOOL CorrectionLumineuseIn,
    char TypeCorrectionLumineuseIn,
    char *NomFichierEtalonIn);

int __stdcall VipsGrimLC13_AttenteNImagesPresentesDansBuffer(TVipsLC13 *TVLC13, unsigned int nombre);   // attente de N images présentes dans le buffer image
int __stdcall VipsGrimLC13_AttenteNFichiersPresentsDansBuffer(TVipsLC13 *TVLC13, unsigned int nombre);  // attente de N fichiers présents dans le buffer fichiers

/********************************************************/

#define CARTE_PULSAR        11                          // gestion carte Pulsar
#define CARTE_DIPIX_LPG     12                          // gestion carte DipixLPG

#define VIPSGRIMLC13_CARTE_VIDEO_INCONNUE_ERR                               622         // type de carte vidéo inconnue
#define VIPSGRIMLC13_CREATION_THREAD_ACQUISITION_DIPIX_LPG_ERR              623         // erreur création thread acquisition DipixLPG
#define VIPSGRIMLC13_LANCEMENT_THREAD_ACQUISITION_DIPIX_LPG_ERR             624         // erreur lancement thread acquisition DipixLPG
#define VIPSGRIMLC13_LIBERATION_HANDLE_THREAD_ACQUISITION_DIPIX_LPG_ERR     625         // erreur libération handle du thread acquisition DipixLPG

// initialisation de la structure TVipsLC13 (gestion carte Pulsar et DipixLPG)
int __stdcall VipsGrimLC13_Init_Acquisition(TVipsLC13 *TVLC13,
    char TypeCarteVideoIn,
    unsigned int NbBuffersDmaIn,
    unsigned int NbBuffersImagesIn,
    unsigned int NbBuffersFichiersIn,
    char *NomFichierCameraIn,
    char *CheminStockageImagesIn,
    char ResolutionImagesIn,
    char TypeFichierIn,
    unsigned int QualiteCompressionECFIn,
    unsigned int QualiteCompressionJPGIn,
    unsigned char SeuilBinarisationTIFFIn,
    char OrientationRectoIn1, char OrientationRectoIn2,
    char OrientationVersoIn1, char OrientationVersoIn2,
    BOOL AcquisitionFenetreeIn,
    unsigned int LargeurFenetreIn, unsigned int HauteurFenetreIn,
    unsigned int OffsetXFenetreIn, unsigned int OffsetYFenetreIn,
    PTR_FonctionUserDataBufferFichiers_STDCALL Ptr_FonctionUserDataBufferFichiersIn_STDCALL,
    unsigned int TailleZoneUserDataBufferFichiersIn,
    BOOL CorrectionLumineuseIn,
    char TypeCorrectionLumineuseIn,
    char *NomFichierEtalonIn);

// génère une image fictive dans le buffer image
int __stdcall VipsGrimLC13_GenereImageFictive(TVipsLC13 *TVLC13);

#ifdef __cplusplus
}
#endif

#endif /* __VIPSGRIM_H__ */