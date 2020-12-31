/*
    filename:   misc.h
    Synopsis:
    Date:       July the 20th 1992

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved.
*/

#ifndef __MISC_H__
#define __MISC_H__

#ifndef __TSERIES_H__
#include <tseries.h>
#endif /* __TSERIES_H__ */

#if defined(WATCOM) || defined(BORLAND_32)
#define INT86(inter,xregs,x1regs,xsregs) int386x(inter,xregs,x1regs,xsregs)
#define XTYPE  w
#else
#define INT86(inter,xregs,x1regs,xsregs) int86x(inter,xregs,x1regs,xsregs)
#define XTYPE  x
#endif

/* memory mapping */
#define  ML_1MEG   1048576L      /* 1 MEG  */
#define  MLBSIZE   4096          /* 4K  main line buffer size */
#define  SEGSIZE   16            /* 64K segment size for tiling */
                                 /* number of 4k blocks */
#define  ML_1K     1024          /* 1 k bytes */
#define  ML_4K     4096          /* 4 k bytes */

/* used by copymem.c  */
#define  ML_DOWN    0            /* vertical direction */
#define  ML_UP      1
#define  ML_RIGHT   0            /* horizontal direction */
#define  ML_LEFT    1

// used by DPMI.c
#define  HMEMSIZE  1024          /* 4M high memory size */
                                 /* number of 4k blocks */
#define  BLOCK     4096          /* 4k block */

// defines for the vertical
// display synchronization

#define DISP_VSYNC_ENABLE   1
#define DISP_VSYNC_DISABLE  0

/* Internal lut control */
#define ML_LUTSIZE         256         /* number of bytes in a LUT */
#define ML_MASK_RED      0x000000FFL
#define ML_MASK_GREEN    0x0000FF00L
#define ML_MASK_BLUE     0x00FF0000L
#define ML_MASK_PALETTE  0x18

/* defines for masks */
#define ML_MASK_MEMTYPE       0x40
#define ML_MASK_LLENGTH       0x30
#define ML_MASK_PIXSIZE       0x80

/* keying defines */
#define ML_KEYCOLOR           0xFF
#define ML_KEYMASK            0x00

/* setregs registers */
#define ML_HFIELD              1
#define ML_ACCDMA              2
#define ML_ACCHOST             3
#define ML_HINTLC              4
#define ML_LLENGTH             5
#define ML_MEMTYPE             6
#define ML_PSIZE               7

/* sysinq registers */
#define ML_MEMORG              0

// mlIO... functions
#define FAST_TRANSFERS         1

// access host configuration model
#define ML_PIXEL              3
#define ML_COLOUR             4

#define ML_ACCPIXEL           0
#define ML_ACCOLOUR           1

// core  functions
#define ML_FILE_READ         0
#define ML_FILE_WRITE        1

// flags used to test mlInit()
#define BOARD_NOT_MAPPED   0xcafe
#define BOARD_MAPPED       0xfade

#ifdef _MSC_VER
#if (_MSC_VER <= 700)
void MFAR* MK_FP(short seg, short ofs);
#endif
#endif

#endif /* __MISC_H__ */

