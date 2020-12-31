#ifndef __VIPSIMAG_H__
#define __VIPSIMAG_H__

#ifdef __cplusplus
extern "C" {
#endif

/******************************************************************/
/*                             LIBRAIRIE                          */
/*----------------------------------------------------------------*/
/* Librairie : vipsimag.dll                                       */
/* Titre : Gestion images BMP, ECF, TIFF G4, JPEG                 */
/* Contenu: - lecture, �criture de fichiers BMP et ECF            */
/*          - cr�ation d'une image BMP                            */
/*          - rotation (90�, 180� et 270�) d'une image BMP        */
/*          - sym�trie horizontale et verticale d'une image BMP   */
/*          - compression des images BMP en ECF                   */
/*          - d�compression des images ECF en BMP                 */
/*          - compression des images BMP en ECF et �criture       */
/*            directe                                             */
/*          - d�coupe d'une image dans une autre et cr�ation d'une*/
/*            nouvelle image                                      */
/*          - auto-suppression des bords noirs de l'image captur�e*/
/*          - capture d'un �talon lumineux sur une image          */
/*          - lecture et �criture d'un �talon lumineux sur le     */
/*            disque                                              */
/*          - correction lumineuse d'une image (� partir de       */
/*            l'�talon et/ou verticale)                           */
/*          - compression des images BMP en JPG et �criture       */
/*            directe                                             */
/*          - d�compression des images JPG en BMP                 */
/*          - compression des images BMP en TIFF G4 et �criture   */
/*            directe                                             */
/*          - d�compression des images TIFF G4 en BMP             */
/* Version : v1.20                                                */
/* D�velopp�e par : M. Tola                                       */
/*----------------------------------------------------------------*/
/* MODIFICATIONS                                                  */
/*----------------------------------------------------------------*/
/* Date de la modification : 30/12/97                             */
/* Code : - LEX                                                   */
/*        - gestion de la d�coupe des images                      */
/*        - gestion de l'�talon                                   */
/*        - gestion de la correction lumineuse d'images           */
/*        - format JPEG                                           */
/*        - format TIFF G4                                        */
/*        - rajout de contr�les suite � la d�compression d'une    */
/*          image ECF en une image BMP dans ECF_ECF2BMP           */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/******************************************************************/
/* Date de la modification : 12/01/98                             */
/* Code : - LEX                                                   */
/*        - am�lioration de la gestion des orientations : ->      */
/*          modifications des fonctions BMP_Init, BMP_Lit,        */
/*          BMP_Orientation, ECF_ECF2BMP, BMP_Decoupe,            */
/*          BMP_JPEG_Disque2BMP, BMP_TIFFG4_Disque2BMP            */
/*          (Cf. variable Inversion).                             */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 19/01/98                             */
/* Code : - LEX                                                   */
/*        - gestion des fonds blancs dans BMP_CoupeBords          */
/*        - gestion des fonds blancs dans Etalon_Capture          */
/*        - rajout du contr�le pour �viter un pixel < 0 dans      */
/*          BMP_CorrectionLuminosite                              */
/*        - rajout du param�tre Orientation dans BMP_Ecrit        */
/*          pour �viter un nouveau retournement de l'image apr�s  */
/*          �criture de celle-ci sur disque                       */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 20/03/98                             */
/* Code : - LEX                                                   */
/*        - chargement dynamique des dlls                         */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 21/04/98                             */
/* Code : - LEX                                                   */
/*        - VipsImag_DechargeDlls modifi�e pour mieux g�rer le    */
/*          d�chargement dynamique des Dlls                       */
/*        - rajout du param�tre Orientation dans                  */
/*          BMP_BMP2JPEG_Disque pour �viter un nouveau            */
/*          retournement de l'image apr�s �criture de celle-ci    */
/*          sur disque                                            */
/*        - rajout du param�tre Orientation dans                  */
/*          BMP_BMP2TIFFG4_Disque pour �viter un nouveau          */
/*          retournement de l'image apr�s �criture de celle-ci    */
/*          sur disque                                            */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 26/05/98                             */
/* Code : - LEX                                                   */
/*        - ajout des fonctions BMP_CalculeHistogramme,           */
/*          BMP_ExpansionDynamique et BMP_RehaussementContraste   */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 16/09/98                             */
/* Code : - LDB                                                   */
/*        - gestion image noire pure dans BMP_ExpansionDynamique  */
/* Code : - LEX                                                   */
/*        - am�lioration SYMETRIE_H et SYMETRIE_V                 */
/*          dans BMP_Orientation                                  */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 17/09/98                             */
/* Code : - LEX                                                   */
/*        - gestion largeur et hauteur de la d�coupe multiple de 4*/
/*                                                                */
/* Modifi� par : M. Tola                                          */
/*----------------------------------------------------------------*/
/* Date de la modification : 18/09/98                             */
/* Code : - LEX                                                   */
/*        - gestion erreur IMAGES_FORMAT_BM_ERR am�lior�e         */
/*          dans BMP_Lit                                          */
/*                                                                */
/* Modifi� par : M. Tola                                          */
/******************************************************************/

// includes de Windows
#include <windows.h>
#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include <io.h>
#include <fcntl.h>
#include <sys/stat.h>
#include <time.h>
#include <malloc.h>

// taille maximale d'une chaine de caract�res
#define TAILLE_MAX_CHAINE             256

// #define des erreurs de la DLL
#define IMAGES_MALLOC_ERR          101        // erreur d'allocation m�moire
#define IMAGES_LOCK_ERR            102        // erreur de 'lock' m�moire
#define IMAGES_FREE_ERR            103        // erreur de lib�ration m�moire

#define IMAGES_OPEN_ERR            111        // erreur d'ouverture fichier
#define IMAGES_WRITE_ERR           112        // erreur d'�criture fichier
#define IMAGES_READ_ERR            113        // erreur de lecture fichier
#define IMAGES_CLOSE_ERR           114        // erreur de fermeture de fichier
#define IMAGES_EOF_ERR             115        // on a atteint la fin de fichier
#define IMAGES_SIZE_ERR            116        // on signale le fait que l'on ne peut pas connaitre la taille du fichier
#define IMAGES_NB_READ_ERR         117        // on n'a pas lu le nb d'octets voulus
#define IMAGES_FORMAT_BM_ERR       118        // le format de l'image ne correspond pas au
                                              // format attendu
#define IMAGES_COMMIT_ERR          119        // erreur de commit

#define IMAGES_ORIENTATION_ERR     121        // erreur d'orientation de l'image

// cf. #define 131, 132, 133 ci-dessous

#define IMAGES_DECOUPE_POS_ERR      141       // mauvaise position de d�coupe
#define IMAGES_DECOUPE_TAILLE_ERR   142       // mauvaise taille de d�coupe
#define IMAGES_ETALON_POS_ERR       143       // mauvaise position de l'�talon
#define IMAGES_ETALON_TAILLE_ERR    144       // mauvaise taille de l'�talon

// #define des op�rations possibles sur les images
#define SYMETRIE_H          11            
#define SYMETRIE_V          12
#define ROTATION_90         13
#define ROTATION_180        14
#define ROTATION_270        15
#define AUCUNE_OPERATION    16
#define ROTATION_270H       17

// #define des diff�rents types de correction de luminosit�
#define CORRECTION_A_PARTIR_ETALON          11
#define CORRECTION_VERTICALE                12
#define TOUTES_CORRECTIONS_POSSIBLES        13

// #define du seuil de blanc valide pour eviter les pixels sombres,
// donc incoh�rents dans une image �talon blanche
#define SEUIL_BLANC_VALIDE                  200

// #define des types de fond
#define FOND_NOIR       1
#define FOND_BLANC      2  

// d�finition de la structure TVipsBMP
// structure permettant la gestion des images BMP
typedef struct
{
    // BITMAPFILEHEADER et BITMAPINFO sont des structures propres � Windows,
    // dans la documentation de Windows, on utilise BmFH comme nom de pointeur
    // pour utiliser la structure BITMAPFILEHEADER et 
    // on utilise BmI comme nom de pointeur
    // pour utiliser la structure BITMAPINFO

    BITMAPFILEHEADER     *BmFH;       // pointeur sur la structure d'ent�te du fichier bitmap
    BITMAPINFO           *BmI;        // pointeur sur la structure d'info du bitmap
    HGLOBAL              Handle;      // handle de la zone m�moire de l'image
    unsigned char        *RawData;    // pointeur sur l'image elle-m�me
    BOOL                 Inversion;   // pour notifier une �ventuelle inversion (x,y) -> (y,x)

    unsigned char        futures[3];  // tableau d'octets et
    unsigned char        *future[2];  // tableau de pointeurs pour "r�server de la place"
                                      // pour des extensions futures et amener la taille
                                      // de la structure � 32 octets
} TVipsBMP;

// d�finition de la structure TVipsInfoEtalon
typedef struct
{
    // largeur et hauteur de l'image �talon
    unsigned int LargeurImageRef, HauteurImageRef;

    // moyenne maximale d'une ligne de l'image �talon
    float MoyenneLigneMax;

    unsigned int PosX, PosY;        // position du rectangle �talon
    unsigned int Largeur, Hauteur;  // largeur et hauteur du rectangle �talon
    float MoyenneRectangle;         // moyenne du rectangle �talon

} TVipsInfoEtalon;

// d�finition de la structure �talon
typedef struct
{
    TVipsInfoEtalon TVIE;       // variable du type TVipsInfoEtalon
    float *MoyennesParLignes;   // pointeur vers le tableau dynamique contenant la moyenne de chaque ligne de l'�talon

} TVipsEtalon;

// Pr�d�claration des fonctions export�es
int __stdcall BMP_Init(TVipsBMP *TVB, int X, int Y);            // Initialise une structure TVipsBMP afin d'accueillir une image BMP non compress�e en 256 niveaux de gris, 1 pixel = 1 octet
int __stdcall BMP_Lit(TVipsBMP *TVB, char *NomFichier);         // Lit � partir du disque une image BMP et stocke celle-ci dans une structure TVipsBMP
int __stdcall BMP_Ecrit(TVipsBMP *TVB, char *NomFichier, BOOL Orientation);       // Ecrit sur disque l'image BMP associ�e � la structure TVipsBMP
int __stdcall BMP_Orientation(TVipsBMP *TVB, int Orientation);  // Effectue une sym�trie horizontale, ou verticale, une rotation de 90�, 180� ou 270 � de l'image
int __stdcall BMP_Fin(TVipsBMP *TVB);                           // Lib�re les ressources associ�es � la structure TVipsBMP

int __stdcall BMP_Decoupe(TVipsBMP *TVB_In, TVipsBMP *TVB_Out, unsigned int PosX, unsigned int PosY,
    unsigned int Largeur, unsigned int Hauteur);    // d�coupe une image dans une autre : cr�e une nouvelle image 
int __stdcall BMP_CoupeBords(TVipsBMP *TVB_In, TVipsBMP *TVB_Out, unsigned char Seuil, char fond); // coupe les bords noirs de l'image
int __stdcall Etalon_Lit(TVipsEtalon *TVE, char *NomFichier);   // lit les infos de l'�talon � partir d'un fichier
int __stdcall Etalon_Capture(TVipsEtalon *TVE, unsigned int PosX, unsigned int PosY,
    unsigned int Largeur, unsigned int Hauteur, TVipsBMP *TVB);     // calcul des infos relatives � l'�talon � partir d'une image 'blanche'
int __stdcall Etalon_Ecrit(TVipsEtalon *TVE, char *NomFichier);     // �crit les infos de l'�talon sur le disque
void __stdcall Etalon_Fin(TVipsEtalon *TVE);    // lib�re la structure TVipsEtalon

void __stdcall BMP_CorrectionLuminosite(TVipsBMP *TVB, TVipsEtalon *TVE, char TypeCorrection);  // corrige l'image � partir d'un �talon

void __stdcall GetDLLVipsImagVersion(char *Version);              // Retourne le num�ro de la version de la DLL sous forme d'une chaine de caract�res
int __stdcall VipsImag_Erreur(unsigned int NumeroErreur, char *ChaineErreur);   // Retourne la chaine de caract�res correspondante au num�ro d'erreur pass� en param�tre

/********************************************************/

// gestion des fichiers ECF

// #define des erreurs de la DLL
#define IMAGES_COMPRESS_ERR        131        // erreur pendant la compression
#define IMAGES_DECOMPRESS_ERR      132        // erreur pendant la d�compression
#define IMAGES_TAUX_ERR            133        // la valeur du taux de compression est <1

// d�finition de la structure TVipsECF
// structure permettant la gestion des images ECF
typedef struct
{
    HGLOBAL              Handle;           // handle de la zone m�moire de l'image
    unsigned char        *CompData;        // pointeur sur l'image compress�e
    unsigned long int    TailleFichier;    // contient la taille du fichier et donc de l'image
    unsigned char        *future[5];       // tableau de pointeurs pour "r�server de la place"
                                           // pour des extensions futures et amener la taille
                                           // de la structure � 32 octets
} TVipsECF;

// Pr�d�claration des fonctions export�es
int __stdcall BMP_BMP2ECF_Disque(TVipsBMP *TVB, char *NomFichier, unsigned int TauxCompression);    // Compresse une image BMP en une image ECF selon un certain taux de compression

int __stdcall ECF_BMP2ECF(TVipsECF *ECF, TVipsBMP *TVB, unsigned int TauxCompression);      // Compresse une image BMP en une image ECF selon un certain taux de compression
int __stdcall ECF_ECF2BMP(TVipsECF *ECF, TVipsBMP *TVB);        // D�compresse une image ECF en une image BMP
int __stdcall ECF_Lit(TVipsECF *ECF, char *NomFichier);         // Lit � partir du disque une image ECF et stocke celle-ci dans une structure TVipsECF
int __stdcall ECF_Ecrit(TVipsECF *ECF, char *NomFichier);       // Ecrit sur disque l'image ECF associ�e � la structure TVipsECF
int __stdcall ECF_Fin(TVipsECF *ECF);                           // Lib�re les ressources associ�es � la structure TVipsECF


/********************************************************/

// gestion des fichiers TIFF et JPEG

// #define des erreurs de la DLL
#define IMAGES_IMPORT_DIB_ERR               151     // probl�me pendant l'importation du DIB
#define IMAGES_CONTROL_SET_ERR              152     // probl�me pendant le positionnement du taux de compression
#define IMAGES_SAVE_FILE_ERR                153     // probl�me pendant la sauvegarde du fichier
#define IMAGES_LOAD_FILE_ERR                154     // erreur pendant la lecture du fichier
#define IMAGES_DIB_PNTR_GET_ERR             155     // erreur pendant l'obtention d'un pointeur sur le DIB
#define IMAGES_DIB_PALETTE_PNTR_GET_ERR     156     // erreur pendant l'obtention d'un pt sur la palette
#define IMAGES_DIB_BITMAP_PNTR_GET_ERR      157     // erreur pendant l'obtention d'un pt sur le bitmap
#define IMAGES_DIB_PIXEL_GET_ERR            158     // erreur pendant l'obtention d'une valeur d'un pixel
#define IMAGES_COLOR_REDUCE_TO_BITONAL_ERR  159     // erreur pendant la binarisation
#define IMAGES_SAVE_FD_ERR                  160     // probl�me pendant la sauvegarde du fichier

#define IMAGES_QUALITE_ERR                  171     // la valeur du taux de compression n'est pas correcte
#define IMAGES_CREATE_FILE_ERR              172     // erreur pendant la cr�ation du fichier
#define IMAGES_CLOSE_HANDLE_ERR             173     // probl�me pendant la fermeture du fichier

// Compresse une image BMP en une image JPEG selon un certain taux de compression (1 � 100). Ecrit l'image JPEG sur le disque.
int __stdcall BMP_BMP2JPEG_Disque(TVipsBMP *TVB, char *NomFichier, unsigned int Qualite, BOOL Orientation);

// Lit � partir du disque une image JPEG et stocke celle-ci dans une structure TVipsBMP
int __stdcall BMP_JPEG_Disque2BMP(TVipsBMP *TVB, char *NomFichier);

// Compresse une image BMP en une image TIFF G4. Ecrit l'image TIFF sur le disque.
// La binarisation se fait par le biais du param�tre Seuil.
int __stdcall BMP_BMP2TIFFG4_Disque(TVipsBMP *TVB, char *NomFichier, unsigned int Seuil, unsigned int Resolution, BOOL Orientation);

// Lit � partir du disque une image TIFF G4 et stocke celle-ci dans une structure TVipsBMP.
int __stdcall BMP_TIFFG4_Disque2BMP(TVipsBMP *TVB, char *NomFichier);

/********************************************************/

#define IMAGES_DLL_NON_TROUVEE_ERR          181     // dll non pr�sente ou non accessible
#define IMAGES_FONCTION_NON_TROUVEE_ERR     182     // fonction non trouv�e � l'int�rieur de la dll
#define IMAGES_DLL_NON_LIBERABLE_ERR        183     // erreur pendant la lib�ration dynamique de la dll       
#define IMAGES_RESOLUTION_SET_ERR           184     // erreur pendant le renseignement de la r�solution de l'image

int __stdcall VipsImag_ChargeFonctionsCets();       // charge dynamiquement les fonctions de la DLL Cets
int __stdcall VipsImag_ChargeFonctionsAccusoft();   // charge dynamiquement les fonctions de la DLL Accusoft
int __stdcall VipsImag_DechargeDlls();              // d�charge dynamiquement les Dlls

/********************************************************/

#define VIPSIMAG_DLL_VIPSPROD_NON_TROUVEE_ERR                               191 // dll VipsProd non trouv�e
#define VIPSIMAG_DLL_FONCTION_AUTORISATION_ACCES_PRODUIT_NON_TROUVEE_ERR    192 // fonction AutorisationAccesProduit non trouv�e dans la DLL VipsProd
#define VIPSIMAG_DLL_FREE_LIBRARY_ERR                                       193 // erreur de lib�ration de la DLL VipsProd
#define VIPSIMAG_DLL_AUTORISATION_ACCES_DLL_REFUSEE_ERR                     194 // autorisation d'acc�s refus�e

// calcule l'histogramme d'une image TVipsBMP et retourne les bornes de celui-ci en fonction du seuil pass� en param�tre
void __stdcall BMP_CalculeHistogramme(TVipsBMP *TVB, unsigned char Seuil, double *BorneInf, double *BorneSup, unsigned int *Histogramme);

// '�tend' l'histogramme de l'image vers une limite aux bornes, c'est � dire vers 0 et 255
void __stdcall BMP_ExpansionDynamique(TVipsBMP *TVB, unsigned char Seuil);

// r�hausse le contraste d'une image par le biais d'un filtre Laplacien
int __stdcall BMP_RehaussementContraste(TVipsBMP *TVB, double Lambda);

#define IMAGES_ACCESS_ERR                                                   195  // fichier non trouv�

// trouve le seuil optimal pour un binarisation future
void __stdcall BMP_CalculeSeuilBinarisationOptimal(TVipsBMP *TVB, unsigned char *SeuilOptimal);
// ecrit sur disque l'image BMP N&B associ�e � la structure TVipsBMP
int __stdcall BMP_Bmp8ToBmp1_Disque(TVipsBMP *TVB, unsigned char Seuil, char *NomFichier);
// lit � partir du disque une image BMP N&B (1 pixel = 1 bit) et stocke celle-ci dans une structure TVipsBMP
int __stdcall BMP_Bmp1_DisqueToBmp8(TVipsBMP *TVB, char *NomFichier);

/********************************************************/

#define IMAGES_L_INIT_BITMAP_ERR                                            196     // erreur fonction L_InitBitmap
#define IMAGES_L_CONVERT_FROM_DIB_ERR                                       197     // erreur fonction L_ConvertFromDIB
#define IMAGES_L_SAVE_BITMAP_ERR                                            198     // erreur fonction L_SaveBitmap
#define IMAGES_L_LOAD_BITMAP_ERR                                            199     // erreur fonction L_LoadBitmap
#define IMAGES_L_CONVERT_TO_DIB_ERR                                         1101    // erreur fonction L_ConvertToDIB

// charge dynamiquement les fonc. des DLLs LeadTools10
int __stdcall VipsImag_ChargeFonctionsLeadTools10();
// compresse une image BMP en une image JPEG selon un certain taux de compression (2 � 255). Ecrit l'image JPEG sur le disque
int __stdcall BMP_BMP2JPEG_DisqueLeadTools10(TVipsBMP *TVB, char *NomFichier, unsigned int Qualite);
// lit � partir du disque une image JPEG et stocke celle-ci dans une structure TVipsBMP
int __stdcall BMP_JPEG_DisqueLeadTools10_2_BMP(TVipsBMP *TVB, char *NomFichier);

// lit � partir du disque une image BMP N&B ou BMP 256
int __stdcall BMP_Lit2(TVipsBMP *TVB, char *NomFichier);
// lit � partir du disque une image JPEG avec Accusoft ou avec LeadTools
int __stdcall JPEG_Lit(TVipsBMP *TVB, char *NomFichier);

/********************************************************/

/********************************************************/
// #define 2000 d�j� pris par IJG 
/********************************************************/

// fonctions relatives au JPEG IJG (The Independent JPEG Group's JPEG software)

// compresse une image BMP en une image JPEG selon un certain taux de compression (1 � 255). Ecrit l'image JPEG sur le disque
int __stdcall BMP_BMP2JPEG_DisqueIJG(TVipsBMP *TVB, char *NomFichier, unsigned int Qualite);
// lit � partir du disque une image JPEG et stocke celle-ci dans une structure TVipsBMP
int __stdcall BMP_JPEG_DisqueIJG_2_BMP(TVipsBMP *TVB, char *NomFichier);

#ifdef __cplusplus
}
#endif

#endif /* __VIPSIMAG_H__ */