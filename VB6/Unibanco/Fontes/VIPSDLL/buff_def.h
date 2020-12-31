/******************************************************************/
/*                           MODULE                               */
/*----------------------------------------------------------------*/
/* Programme principal : buffers.cpp                              */
/* Contenu du module : gestion buffers images et buffers fichiers */
/* Développé par : M. Tola                                        */
/*----------------------------------------------------------------*/
/* MODIFICATIONS                                                  */
/*----------------------------------------------------------------*/
/* Date de la modification : Objet de la modification             */
/* Code : Identification de la modification                       */
/* Modifié par : Nom du développeur                               */
/******************************************************************/

#ifndef __BUFF_DEF_H__
#define __BUFF_DEF_H__

#ifdef __cplusplus
extern "C" {
#endif

// includes de Windows
#include <windows.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>

// include de "vipsimag.dll" et "vipsgrab.dll"
#include "vipsimag.h"
#include "vipsgrab.h"

// nombre maximal de buffers images autorisés
#define NB_MAX_BUFFERS_IMAGES                40

// #define des erreurs
#define BUFFER_IMAGES_ALLOC_ERR                           401       // erreur d'allocation mémoire
#define BUFFER_IMAGES_LOCK_ERR                            402       // erreur de lock mémoire
#define BUFFER_IMAGES_UNLOCK_ERR                          403       // erreur de unlock mémoire
#define BUFFER_IMAGES_FREE_ERR                            404       // erreur de libération mémoire

#define BUFFER_IMAGES_TROP_DE_BUFFERS_DEMANDES_ERR        411       // trop de buffers images demandés par rapport au nombre autorisé
#define BUFFER_IMAGES_PLEIN_ERR                           412       // buffer images plein
#define BUFFER_IMAGES_VIDE_ERR                            413       // buffer images vide

// définition de la structure TVipsBufferImages
// structure permettant la gestion des buffers images
typedef struct
{
    unsigned int NbBuffers;             // nb de buffers images utilisés
    unsigned int TailleBuffer;          // taille d'un buffer image
        
    TVipsBMP *TVB[NB_MAX_BUFFERS_IMAGES];           // buffer image, c'est une zone TVipsBMP
    HGLOBAL HandleBuffer[NB_MAX_BUFFERS_IMAGES];    // handle mémoire de la zone ci-dessus

    unsigned int NbElemsDansBuffer;     // nb d'éléments dans le buffer
    unsigned int PointeurInBuffer;      // pointeur d'entrée du buffer
    unsigned int PointeurOutBuffer;     // pointeur de sortie du buffer

    unsigned int LargeurImage, HauteurImage;        // largeur et hauteur de l'image dans le buffer

} TVipsBufferImages;

//*****************************************************************

// nombre maximal de buffers fichiers autorisés
#define NB_MAX_BUFFERS_FICHIERS                40

// taille maximale d'une chaine de caractères
#define TAILLE_MAX_CHAINE                      256

#define BUFFER_FICHIERS_TROP_DE_BUFFERS_DEMANDES_ERR      421       // trop de buffers fichiers demandés par rapport au nombre autorisé
#define BUFFER_FICHIERS_PLEIN_ERR                         422       // buffer fichiers plein
#define BUFFER_FICHIERS_VIDE_ERR                          423       // buffer fichiers vide

#define BUFFER_FICHIER_ALLOC_ERR                          431       // erreur d'allocation du buffer fichiers

typedef void USERDATA_BUFFER_FICHIERS;      // création d'un type USERDATA_BUFFER_FICHIERS

// définition de la structure TVipsBufferFichiers
// structure permettant la gestion des buffers fichiers
typedef struct
{
    unsigned int NbBuffers;     // nb de buffers fichiers utilisés
    
    unsigned char InformationFace[NB_MAX_BUFFERS_FICHIERS];      // information sur la face du fichier
    char *NomFichier[NB_MAX_BUFFERS_FICHIERS];                   // nom du fichier

    unsigned int NbElemsDansBuffer;     // nb d'éléments dans le buffer
    unsigned int PointeurInBuffer;      // pointeur d'entrée du buffer
    unsigned int PointeurOutBuffer;     // pointeur de sortie du buffer

    USERDATA_BUFFER_FICHIERS *PointeurUserData[NB_MAX_BUFFERS_FICHIERS];             // pointeur vers une structure 'UserData'
    unsigned int TailleZoneUserData;    // taille de la zone UserData

} TVipsBufferFichiers;

#ifdef __cplusplus
}
#endif

#endif /* __BUFF_DEF_H__ */
