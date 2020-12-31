/******************************************************************/
/*                             LIBRAIRIE                          */
/*----------------------------------------------------------------*/
/* Librairie : vipsdrv.h						                  */
/* Titre : Pilotage des lecteurs LA93 ou LG90                     */
/* Contenu: Procedures necessaires au traitement de documents     */
/* Version : v3.02								                  */
/* Développée par : Myriam RIFFARD		                          */
/*																  */
/* L'appel de la fonction vipsGetLibelleStatus nécessite 2 fichiers*/
/* texte : Erreur.txt et ErrCom.txt.						      */
/* Tous les codes Retour d'erreur peuvent prendre la valeur des   */
/* erreurs contenues dans le fichier Erreur.txt                   */
/*----------------------------------------------------------------*/
/* MODIFICATIONS                                                  */
/*----------------------------------------------------------------*/
/* Date de la modification :  19 fevrier 1998                     */
/* Code : LDV                                                     */
/*        Integration de la nouvelle serie (Serie.DLL)            */
/*        Suppression des ifndef NONDOS                           */
/*        Ajout de 3 statuts d'erreurs (-105,-106,-107)           */
/*        Rajout de 3 variables globales                          */
/*        Modification de 6 procedures                            */  
/*        Modification du type du parametre ListeCom              */             
/* Modifié par : Myriam RIFFARD                                   */
/******************************************************************/
/* Date de la modification :  11 mai 1998  (version 3.10)         */
/* Code : LDB                                                     */
/*        ajout de desallocation de memoire                       */  
/*        creation de 3 nouvelles constantes                      */
/*          VER_PB_ACTIVATION_TIMEOUT_LECTURE  -112               */
/*          VER_PB_DESACTIVER_TIMEOUT          -113               */
/*          TIMEOUT_ECRITURE     5000                             */ 
/*        modification du fichier erreur.txt                      */   
/*        Modification du fichier espion, mise en place d'un      */
/*         espion global                                          */   
/*        Ajout du TimeOut en lecture sur la com                  */ 
/*        Ajout de constantes pour le fichier espion et de 2      */
/*        messages d'erreur                                       */
/*          NOMBRE_BUFFERS_ESPION  10                             */
/*          TAILLE_BUFFER_ESPION  255                             */
/*          VER_PB_ACTIVE_ESPION -114                             */
/*          VER_PB_DESACTIVE_ESPION -115                          */
/*       Ajout d'une fonction d'interrogation de la EPROM         */
/*       Ajout d'une fonction d'analyse du lecteur connecte :     */
/*              vipsTypeLecteur                                   */
/*        Ajout de d'un message d'erreur                          */
/*          PB_LECTEUR_INCONNUE    -116                           */ 
/* Modifié par : Myriam RIFFARD                                   */
/******************************************************************/
/* Date de la modification :  20 juillet 1998  (version 3.11A)    */
/*                             6 aout    1998                     */
/* Creation de la version 3.12                                    */
/* Code : LDV                                                     */
/*					                                              */
/*	       NOMBRE_BUFFERS_ESPION  20                              */
/*         Ajout d'élément dans le fichier espion DLL             */
/*         Modification de vipsSortDoc : ajout de complément      */
/*             aux erreurs de cases (-101) et d'impression (-102) */
/*         Modification de vipsCommand : ajout des erreurs -101   */
/*             et -102                                            */
/*         Correction d'un bug pour le renvoi des erreurs VER_SYNTAX*/
/*         Rajout de la fonction vipsEnvoiFinger                  */
/*         Rajout de fonctions pour le traitement de l'image      */
/*           - vipsConstrMotInitImageStdCall                      */
/*           - vipsAddStationImageStdCall                         */
/*           - vipsArretAcquisitionStdCall                        */
/*           - vipsMarcheAcquisitionStdCall                       */
/*           - vipsAllumeNeonStdCall                              */
/*           - vipsEnvoiNomFichierImageStdCall                    */
/*           - vipsEnvoiFingerLa93StdCall                         */
/*           - vipsEvacuationImageStdCall                         */
/*           - vipsInitMarcheStationImageStdCall                  */
/*     	     - vipsEnvoiInitImageStdCall                          */
/*      	 - vipsRenvoiMotInitLecteurStdCall                    */
/*					                                              */
/* Modifié par : Myriam RIFFARD                                   */
/******************************************************************/

#ifndef __VIPSDRV_H__
#define __VIPSDRV_H__

#ifdef __cplusplus
	extern "C" {
#endif

/* Constantes d'Erreur */

#define VER_NOERROR		 0	 /* Pas d'erreur detectee*/
#define VER_TIMEOUT		-1	 /* Delai d'attente ecoule */
#define VER_EMPTY		-2	 /* Depileur vide */
#define VER_NOTINIT		-3	 /* L'appel a vipsOpen n'a pas ete effectue */ 
#define VER_SYNTAX		-4	 /* Erreur de syntaxe dans une commande dans une
							    reponse ou dans un champ de la communication */
#define VER_BOXFULL		-5	 /* Case de reception pleine */ 
#define VER_BOXEMPTY	-6	 /* Case de reception vide */
#define VER_TIMEOUT_MODE_COMMANDE -7 /* TimeOut du mode Commande du LG90 est incorrecte */
#define VER_COMM		-10	 /* Commande inconnue */
#define	VER_INIT		-30	 /* Mot d'initialisation incorrect */
#define VER_CMDBOX		-40	 /* Commande de traitement de document incorrecte */
#define VER_DOUBLE		-50	 /* Documents introduit en double */
#define VER_JAM			-51	 /* Bourrage dans le parcours du document */
#define VER_FEED		-55	 /* Document depile par erreur */
#define	VER_DOCEXPECTED	-59	 /* Document attendu ,à un endroit du systeme, non vu */
#define VER_LATEDOC		-60	 /* Document en retard */
#define VER_FASTDOC		-61	 /* Document en avance */
#define	VER_WAITINGBCMD	-62	 /* Ordre de case attendu */
#define VER_WAITINGECMD	-63	 /* Ordre d'endossage attendu */
#define	VER_WAITINGPCMD	-64	 /* Ordre d'impression attendu */
#define VER_ANS			-65	 /* Pas de reponse dans l'unite */
#define VER_WAITINGINIT	-66	 /* Ordre d'initialisation attendu */
#define	VER_BOXSTACK	-67	 /* La pile d'ordre de case de l'unite est pleine */
#define	VER_STATUSUNK	-68  /* Statut recu de l'unite inconnu */
#define VER_OUTOFRIBBON	-69	 /* Probleme d'enroulement du ruban du post marqueur */
#define VER_ABS_TIROIR_ECLAIRAGE -80 /*Absence du tiroir d'eclairage*/
#define VER_ABS_LUMIERE -81 /*Absence de lumiere*/
#define VER_SUPERPOSITION -82 /*Superposition des fenetres d'acquisition*/
#define VER_PB_CODEUR -90 /* Probleme codeur dans un module*/
#define	VER_LOADERR		-100 /* DLL deja cherchee; IPL du systeme*/ 
#define	VER_NBRCASE		-101 /* Numéro de case incorrecte */
#define	VER_CARIMP		-102 /* Caractère non imprimable */
#define VER_PB_BUFFER	-103 /* Probleme dans le buffer tournant gerant la communication */
#define VER_FLUSH		-104 /* Erreur lors de l'ecriture directe disque */ 
#define VER_PB_LECTURE_COM    -105 /*Probleme lors de la lecture de la com*/
#define VER_PB_ECRITURE_COM   -106 /*Probleme lors de l'ecriture de com*/
#define VER_PB_FERMETURE_COM  -107 /*Probleme lors de la fermeture de la com*/
#define VER_PB_LOAD_VIPSPRODDLL -108 /*Problème lors du chargement de la fonction appartenant à VipsProd.DLL*/
#define VER_PB_ADRESS_VIPSPRODDLL -109 /*Problème dans l'adresse de la fonction appartenant à VipsProd.DLL*/
#define VER_PB_NUMERO_SERIE -110 /*Numéro de série non valide*/
#define VER_PB_FERMETURE_VIPSPRODDLL -111 /*Problème lors du dechargement de VipsProd.DLL*/
#define VER_PB_ACTIVATION_TIMEOUT_LECTURE  -112 /*Probleme lors de la mise en place du TimeOut en lecture*/
#define VER_PB_DESACTIVER_TIMEOUT -113 /*Probleme lors de la desactivation du TimeOut en lecture*/		
#define VER_PB_ACTIVE_ESPION -114 /*Probleme lors de l'activation du fichier espion*/
#define VER_PB_DESACTIVE_ESPION -115 /*Probleme lors de la desactivation du fichier espion*/
#define PB_LECTEUR_INCONNUE    -116 /*Lecteur inconnue lors de la demande de version d'eprom*/
#define VER_ERREUR_IMAGE  -117 /*Erreur de la station Image*/
#define ERREUR_TYPE_IMAGE -118 /*Erreur dans la définition du type de l'image*/
#define ERREUR_ALLOCATION_MEMOIRE -119 /*Erreur lors de l'allocation mémoire*/
#define ERREUR_CARACTERE_DE_DEBUT_NON_TROUVE  -120 /*Caractere de début non trouvé*/
#define VER_ERREUR_COM_IMAGE   -121 /*Erreur lors de l'utilisation de la com de l'image*/
#define VER_PB_RELEASE_SEMAPHORE -122 /*Erreur de release du semaphore*/
#define VER_ERREUR_WAIT_FAILED  -123 /*Erreur Wait Failed*/
#define VER_ERREUR_WAIT_ABANDONNED  -124 /*Erreur Wait Abandonned*/
#define VER_ERREUR_WAIT_TIMEOUT  -125 /*Erreur Wait Timeout*/
#define VER_ERREUR_REPONSE_WAIT_OBJECT -126 /*Erreur inconnu du WaitForSingleObject*/


/* Port de communications */

#define COM1	1			 
#define COM2	2
#define COM3	3
#define COM4	4

/* Parametres de parite */

#define NoParity	0		/* Pas de parite */
#define EvenParity	1		/* Parite Paire */
#define OddParity	2		/* Parite Impaire */

#define Paire 2
#define Impaire 1 

/* Parametres type de lecteur */

#define vipsLA93 1
#define vipsLG90 2
#define vipsLG91 3

/* Paramèteres type de la station image */

#define ImageRecto 1
#define ImageVerso 2

/*Type de lecteur connecte*/
#define LA93  1
#define LG90  2
#define LG91  3 
#define MC93  4


/* Parametres type de HARDWARE pour le champ hwType de la structure vipsHARDWARE */

#define	vipsBoxes			1	/* Nombre de cases pour le tri */
#define	vipsExtModule		2	/* Presence d'un module d'extention */
#define	vipsEndorser		3	/* Presence d'endosseur */
#define	vipsPrinter			4	/* Presence d'une Imprimante */
#define	vipsMagnPrinter		5	/* Presence d'un postmarquer */	
#define	vipsLenFeelers		6	/* Seuil de longueur pour la detection de double */
#define	vipsFeelers			7	/* seuil de detection pour le double */

/* Parametres type de commande pour le champ cmdType de la structure vipsCOMMAND */

#define vipsBox			1		/* Case de reception */
#define vipsPrint		2		/* Chaine à imprimer */
#define vipsMagnPrint	3		/* Chaine à postmarquer */
#define vipsEndorse		4		/* Endossage */
#define vipsOTHER		5		/* Autre materiel non encore developpe */

/* Parametres booleen */

#define	FALSE	0
#define	TRUE	1

/* TimeOut en lecture sur la serie */

#define TIMEOUT_ECRITURE     5000

/*Constantes pour l'activation du fichier espion global */

#define NOMBRE_BUFFERS_ESPION  20
#define TAILLE_BUFFER_ESPION  255

/*Taille max de la chaine renvoyee par le lecteur pour version des eprom*/
#define TAILLE_CHAINE_VERSION   255

/*Taille des mallocs*/
#define MALLOC_MOT_INIT               80    
#define MALLOC_CHAINE_LUE             250  
#define MALLOC_CHAINE_COURTE          100 
#define MALLOC_TRES_FAIBLE            30    
#define MALLOC_VERSION_MODULE         500  
#define MALLOC_NOM_ET_CHEMIN_FICHIER  400
#define MALLOC_IMPRESSION             100  
#define MALLOC_MESSAGE                200

/*Nombre max de caracteres à lire sur la com*/
#define NB_MAX_CARACTERE_A_LIRE     250

//Timeout pour la reponse à l'init de la station image
#define TIMEOUT_IMAGE  5000

/* Declaration des structures

/* Structure necessaire à la gestion HARDWARE : type de materiel et etat du materiel */
			
typedef struct {
	int hwType;
	int	hwState;
} vipsHARDWARE;

/* Structure necessaire à la gestion des commandes */ 

typedef struct {
	int cmdType;	/* Type de materiel pour lequel la commande est envoyee */ 
	char * Cmd;		/* Commande en elle meme */
} vipsCOMMAND;

/* Parametre de la communication */ 

typedef struct vipsDcb {
	int dcbSize;				/* Taille de la structure */
	unsigned char Type;			/* Type de lecteur connecte */
	unsigned char Port;			/* Numero du port de communications à utiliser */
	unsigned int BaudRate;		/* Vitesse de transmission */
	unsigned char ByteSize;		/* Taille des caracteres dans la transmission */
	unsigned char Parity;		/* Parite */
	unsigned char StopBits;		/* Nombre de bits de stop */
	unsigned int vipsTimeout;	/* Duree du Timeout en millisecondes lors de l'attente d'une lecture */
	char EvtChar;				/* Caractere generant la fin se sequence */
	unsigned int PortAddres;	/* Adresse du port de communication */
	unsigned char PortIRQ;		/* Numero d'IRQ du port de communication */
} vipsDCB;


/* Ouverture de la communication avec le lecteur VIPS */
int __stdcall vipsOpenStdCall(vipsDCB far *lpDCB);

/* Procedure verifiant les parametres de la communication */
/* et initialisant la communication avec le systeme */
int __stdcall vipsOpen2StdCall(vipsDCB far *lpDCB);

/* Fermeture de la communication avec le lecteur VIPS */
void __stdcall vipsCloseStdCall( void );

/*	Demande de lecture d'un document */ 
int __stdcall vipsReadCMC7StdCall(char far *lpBuf,int nSize);

/*	Lecture de la seconde station de lecture */
int __stdcall vipsReadSecondStdCall(char far *lpBuf,int nSize);

/* Permet une lecture simultanee sur les deux stations de lecture */
int __stdcall vipsReadDoubleStdCall(char far *lpBuf1,int nSize1, char far *lpBuf2,int nSize2);

/* Ejecte le dernier document present dans le transport en cas de bourrage ou de depileur vide */
int __stdcall vipsEjectStdCall( void );

/* Demande le statut d'erreur de la machine */
int __stdcall vipsStatusStdCall( char far *lpComp );

/* Renseignement du driver sur le type de materiel pilote */
int __stdcall vipsSetHardwareStdCall( int, vipsHARDWARE far *);

/* Suppression de materiel pilote par le driver */
void __stdcall vipsGetHardwareStdCall( int *, vipsHARDWARE far *);

/* Envoi une commande de traitement pour le document se trouvant la chambre de depression */
int __stdcall vipsCommandStdCall(int, vipsCOMMAND far *);

/* Mise en route des moteurs */
int __stdcall vipsBeginStdCall( void );

/* Arret des machines */ 
int __stdcall vipsEndStdCall( void );

/* Demande du numero de version du driver */
unsigned int __stdcall vipsVersionStdCall( void );

/* Ajoute une seconde station de lecture au pilote */
void __stdcall vipsAddStationStdCall( char *);

/* Creation d'un fichier d'espionnage des communications*/
FILE* __stdcall vipsSetupSpyFileStdCall( char * );

/* Fermeture du fichier espion */
void __stdcall vipsCloseSpyFileStdCall( void );

/* Envoi au materiel de l'ordre de traitement pour le document */
int __stdcall vipsSortDocStdCall( short , char *, char *, int );

/* Defini la marche à suivre apres une evacuation */
void __stdcall vipsSetEjectStatusStdCall(int FlagEject);

/* Initialise le lecteur */
void __stdcall vipsInitStdCall( void );

/* Fermeture du fichier espion */
void __stdcall vipsCloseDLLSpyFileStdCall(void);

/* Creation d'un fichier d'espionnage des appels à Vipsdrv.dll */
FILE* __stdcall vipsSetDLLSpyFileStdCall( char *Fichier );

/* Demande du numero de version de la DLL */
void __stdcall vipsDLLversionStdCall(char * s);

/* Renvoie le libelle exacte de l'etat du systeme */
/* Cette fonction apelle vipsStatus */
char* __stdcall vipsGetLibelleStatusStdCall(int CodeRetour);

/*Recupere le libelle du statut d'erreur (cette fonction est la même que vipsGetLibelleStatus
  mais elle est plus propre)*/
int __stdcall vipsRecupereStatusLibelleStdCall(char* Libelle,int *CodeRetour);

/* Lecture d'une carte */
int __stdcall vipsReadCarteStdCall(char far *lpBuf,int nSize);

/* Passage au mode Commande pour un LG90 */
int __stdcall vipsPassageModeCommandeStdCall(void);

/*Passage au mode Autonome pour un LG90 */
int __stdcall vipsPassageModeAutonomeStdCall(void);

/*Construction du MotInit de la deuxième station de lecture de type OCR*/
void __stdcall vipsMotInitOCRStdCall(unsigned char Unite,
					int DistBordDroitDocChaine,
					int LongueurChaineALire,
					int TypeFonte,
					BOOL ActiveDetectionBlanc,
					char* MotInit);

/*Construction du MotInit de la deuxième station de lecture de type Code à barre*/
void __stdcall vipsMotInitCodeBarreStdCall(int TypeCodeBarre,
						 int SensLecture,
						 int DistanceCodeBarreBordDroitDoc,
						 int LongueurCodeBarre,
						 char* MotInit);

/*Construction du mot d'init pour la station image*/
void __stdcall vipsConstrMotInitIMAGEStdCall(
							  unsigned char TypeFichier,
							  int NbreBufferImage,
							  char* NomFichierDCF,
							  int QualiteCompression,
							  unsigned char RectoVerso,
							  int NbreRetardTolereCompression,							  
							  char* MotInit);

/*Demande de version aux EPROM*/
int __stdcall vipsVersionEPROMStdCall(char far *lpBuf,int nSize);

/*Demande du type de lecteur connecte*/
int __stdcall vipsTypeLecteurStdCall(int* TypeLecteur);

/*Envoi un finger au module image (pour un MC13)*/
int __stdcall vipsEnvoiFingerStdCall( void );

/*Ajout d'une station Image*/
int __stdcall vipsAddStationImageStdCall(vipsDCB far *lpDCBImage,char *MotInitImage);

/*Envoi l'ordre d'arret d'acquisition*/
int __stdcall vipsArretAcquisitionStdCall(BOOL Recto);

/*Envoi l'ordre d'acquisition*/
int __stdcall vipsMarcheAcquisitionStdCall(BOOL Recto);

/*Allume les caméras du module image*/
int __stdcall vipsAllumeNeonStdCall(BOOL AllumeRecto,
								 int ResolCodeur,
								 int DistOuverRecto,
								 int DureeFenAcquiRecto,
								 BOOL AllumeVerso,
								 int DistOuverVerso,
								 int DureeFenAcquiVerso);

/*Envoi a la station image le nom du fichier pour la sauvegarde de l'image*/
int __stdcall vipsEnvoiNomFichierImageStdCall(BOOL Recto,BOOL CapteImage,char* NomFichierImage);

/*Envoi un finger au module image (pour un LA93)*/
int __stdcall vipsEnvoiFingerLA93StdCall(void);

/*Envoi l'ordre d'evacuation au module image*/
int __stdcall vipsEvacuationImageStdCall(BOOL Recto);

/*Envoi l'init et la marche acquisition au module image*/
int __stdcall vipsInitMarcheStationImageStdCall(BOOL Recto,char* MotInitImage);

/*Envoi l'init au module image*/
int __stdcall vipsEnvoiInitImageStdCall(BOOL Recto);

/*Donne le Mot d'init du lecteur*/
void __stdcall vipsRenvoiMotInitLecteurStdCall(char* MotInitLecteur);

#ifdef __OS2__
	#define far
#endif

#ifdef WIN32
	#define far
	#define _export 
	#ifndef pascal
		#define pascal
	#endif
#endif

#ifdef __WIN32__
	#define far
#endif





/* Declaration des fonctions et procedures

/* Ouverture de la communication avec le lecteur VIPS */
int far pascal _export vipsOpen(vipsDCB far *lpDCB);

/* Procedure verifiant les parametres de la communication */
/* et initialisant la communication avec le systeme */
int far pascal _export vipsOpen2(vipsDCB far *lpDCB);

/* Fermeture de la communication avec le lecteur VIPS */
void far pascal _export vipsClose( void );

/*	Demande de lecture d'un document */ 
int far pascal _export vipsReadCMC7(char far *lpBuf,int nSize);

/*	Lecture de la seconde station de lecture */
int far pascal _export vipsReadSecond(char far *lpBuf,int nSize);

/* Permet une lecture simultanee sur les deux stations de lecture */
int far pascal _export vipsReadDouble(char far *lpBuf1,int nSize1, char far *lpBuf2,int nSize2);

/* Ejecte le dernier document present dans le transport en cas de bourrage ou de depileur vide */
int far pascal _export vipsEject( void );

/* Demande le statut d'erreur de la machine */
int far pascal _export vipsStatus( char far *lpComp );

/* Renseignement du driver sur le type de materiel pilote */
int far pascal _export vipsSetHardware( int, vipsHARDWARE far *);

/* Suppression de materiel pilote par le driver */
void far pascal _export vipsGetHardware( int *, vipsHARDWARE far *);

/* Envoi une commande de traitement pour le document se trouvant la chambre de depression */
int far pascal _export vipsCommand(int, vipsCOMMAND far *);

/* Mise en route des moteurs */
int far pascal _export vipsBegin( void );

/* Arret des machines */ 
int far pascal _export vipsEnd( void );

/* Demande du numero de version du driver */
unsigned int far pascal _export vipsVersion( void );

/* Ajoute une seconde station de lecture au pilote */
void far pascal _export vipsAddStation( char *);

/* Creation d'un fichier d'espionnage des communications*/
FILE* far pascal _export vipsSetupSpyFile( char * );

/* Fermeture du fichier espion */
void far pascal _export vipsCloseSpyFile( void );

/* Envoi au materiel de l'ordre de traitement pour le document */
int far pascal _export vipsSortDoc( short , char *, char *, int );

/* Defini la marche à suivre apres une evacuation */
void far pascal _export vipsSetEjectStatus(int FlagEject);

/* Initialise le lecteur */
void far pascal _export vipsInit( void );

/* Fermeture du fichier espion */
void far pascal _export vipsCloseDLLSpyFile(void);

/* Creation d'un fichier d'espionnage des appels à Vipsdrv.dll */
FILE* far pascal _export vipsSetDLLSpyFile( char *Fichier );

/* Demande du numero de version de la DLL */
void far pascal _export vipsDLLversion(char * s);

/* Renvoie le libelle exacte de l'etat du systeme */
/* Cette fonction apelle vipsStatus */
char* far pascal _export vipsGetLibelleStatus(int CodeRetour);

/*Recupere le libelle du statut d'erreur (cette fonction est la même que vipsGetLibelleStatus
  mais elle est plus propre)*/
int far pascal _export vipsRecupereStatusLibelle(char* Libelle,int *CodeRetour);

/* Lecture d'une carte */
int far pascal _export vipsReadCarte(char far *lpBuf,int nSize);

/* Passage au mode Commande pour un LG90 */
int far pascal _export vipsPassageModeCommande(void);

/*Passage au mode Autonome pour un LG90 */
int far pascal _export vipsPassageModeAutonome(void);

/*Construction du MotInit de la deuxième station de lecture de type OCR*/
void far pascal _export vipsMotInitOCR(unsigned char Unite,
					int DistBordDroitDocChaine,
					int LongueurChaineALire,
					int TypeFonte,
					BOOL ActiveDetectionBlanc,
					char* MotInit);

/*Construction du MotInit de la deuxième station de lecture de type Code à barre*/
void far pascal _export vipsMotInitCodeBarre(int TypeCodeBarre,
						 int SensLecture,
						 int DistanceCodeBarreBordDroitDoc,
						 int LongueurCodeBarre,
						 char* MotInit);

/*Demande de version aux EPROM*/
int far pascal _export vipsVersionEPROM(char far *lpBuf,int nSize);

/*Demande du type de lecteur connecte*/
int far pascal _export vipsTypeLecteur(int TypeLecteur);

/*Envoi un finger au module image (pour un MC13)*/
int far pascal _export vipsEnvoiFinger( void );






#ifdef __cplusplus
	}
#endif

#endif /* __VIPSDRV_H__ */

