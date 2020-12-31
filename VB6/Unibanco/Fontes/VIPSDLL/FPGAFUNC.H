#include <string.h>

#ifdef __unix
#include <stdlib.h>
#include <malloc.h>
#include "doslib.h"
#else
#include <alloc.h>
#endif

/* Some types defined here are used with prototypes */
#ifndef __EXEC2_H__     /* Don't include the standard library includes in */
                        /* MATROX internal test programs                  */
#include <stdlib.h>
#include <stdio.h>
#include <ctype.h>
#endif

#ifdef DOSX286
#include <magicdev.h>
#include <phapi.h>
#else
#ifdef __unix
#include <magicdev.h>
#else
#include <dos.h>
#define MA_MS_DOS
#include <magicdev.h>
#endif
#endif


unsigned long LdFPGA(char *, char*, char, int);
unsigned long LdFPGAEISA(char *, char *, char);
void LdFPGAPrtLoading(char *filename);

#define LOAD_FPGA       0
#define CREATE_OUT      1
#define CREATE_CRC      2

#define MAX_LEN_PARAM   128

#define ERRGENE1        0x00000001  
#define ERRGENE2        0x00000002   
#define ERROPTION1      0x00000004
#define ERROPTION2      0x00000008  
#define ERRSALL         0x00000010
#define ERRMALLOC       0x00000020
#define ERRFOPEN1       0x00000040
#define ERRFOPEN2       0x00000080
#define ERRFWRITE       0x00000100
#define ERRFOPEN3       0x00000200
#define ERRCMDOPT       0x00000400        /* Obsolete */
#define ERRNOBRD        0x00000800
#define ERRCRC          0x00001000

#define ERRLOC_0        0x00010000
#define ERRLOC_1        0x00020000
#define ERRLOC_2        0x00040000
#define ERRLOC_3        0x00080000
#define ERRLOC_4        0x00100000
#define ERRLOC_5        0x00200000
#define ERRLOC_6        0x00400000
#define ERRLOC_7        0x00800000
#define ERRLOC_8        0x01000000
#define ERRLOC_9        0x02000000
#define ERRLOC_a        0x04000000
#define ERRLOC_b        0x08000000
#define ERRLOC_c        0x10000000
#define ERRLOC_d        0x20000000
#define ERRLOC_e        0x40000000
#define ERRLOC_f        0x80000000
