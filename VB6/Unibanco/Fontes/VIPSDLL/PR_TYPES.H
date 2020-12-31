/*
    filename:   pr_types.h
    Synopsis:   Include file for image processing library 
                type definitions.
    Date:       18/09/92

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved
*/

#ifndef __PR_TYPES_H__
#define __PR_TYPES_H__

#define  PR_CHAR               char	    
#define  PR_UCHAR     unsigned char	  
#define  PR_SHORT              short	 
#define  PR_USHORT    unsigned short	 
#define  PR_INT                int	 
#define  PR_UINT      unsigned int	 
#define  PR_LONG               long	 
#define  PR_ULONG     unsigned long	  
#define  PR_DOUBLE             double	 
#define  PR_FLOAT              float	 
#define  PR_VOID               void	 
#ifndef WATCOM
#define  PR_MHUGE              huge
#else
#define  PR_MHUGE
#endif
#define  PR_IMPORT             extern           
#define  PR_LOCAL              static           
#define  PR_FAST               register

/* Define ANSI for ANSI standard. Most unixes will require that ANSI */
/* not be defined.                                                   */

#ifndef ANSI
#ifdef __STDC__
#define ANSI
#endif /* __STDC__ */
#endif /* ANSI */

#endif /* __PR_TYPES_H__ */
