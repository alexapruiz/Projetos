/*                                                                                                                                /*
    filename :  WIN_NT.h
    Synopsis:
    Date:       1994/08/28

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved
*/

#ifndef __WIN_NT_H__
#define __WIN_NT_H__

/* define operating system */
#ifndef MSC_NT
#define MSC_NT
#endif

#define MFAR
#define MHUGE
#define FTYPE _stdcall
#define PTYPE
#define MFAR_1M

#undef  far
#define far
#undef  huge
#define huge
#undef  pascal
#define pascal

/* define number of boards */
#define T_N_OF_BOARDS  4

#endif /* __WIN_NT_H__ */


