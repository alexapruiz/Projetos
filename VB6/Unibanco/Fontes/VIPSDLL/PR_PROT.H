/*
    filename:   pr_prot.h
    Synopsis:   Include file for image processing library specific
                type definitions. To be included by the users.
    Date:       18/09/92

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved
*/

#ifndef __PR_PROT_H__
#define __PR_PROT_H__

#include <pr_types.h>

/* Should be defined in pr_sys.h */
#ifndef ANSI
#define ANSI
#endif /* ANSI */

/* Define for surface type */
#define ML_IMAGE_SYS_RAM 	 0x04

/* Func_type to be applied to the resulting pixel after convolution. */ 
/* Used with Conv3x3() and Conv3x3Mix()                              */
#define ML_NORMAL_CONV         0
#define ML_DIV_CONV            1
#define ML_TRONC_CONV          2
#define ML_CLIP_CONV           3
#define ML_TRUNC_CONV          ML_TRONC_CONV	// To relect the manual



/* The Operation to be applied to the source and destination images.  */
/* Used with MixImage()                                               */
#define ML_ADD_IMAGE           0
#define ML_SUB_TO_IMAGE        1
#define ML_SUB_FROM_IMAGE      2
#define ML_OR_IMAGE            3
#define ML_AND_IMAGE           4
#define ML_XOR_IMAGE           5
#define ML_MIN_IMAGE           6
#define ML_MAX_IMAGE           7

// Ops on image buf and constant
#define ML_ADD_CONST	          8  
#define ML_SUB_CONST	          9  
#define ML_CONST_SUB	         10 
#define ML_AND_CONST	         11 
#define ML_NAND_CONST	      12 
#define ML_OR_CONST	         13 
#define ML_XOR_CONST	         14 
#define ML_NOR_CONST	         15 
#define ML_XNOR_CONST	      16 
#define ML_MIN_CONST	         17 
#define ML_MAX_CONST	    	   18
#define ML_MULT_CONST	 	   19
#define ML_DIV_CONST	    	   20
#define ML_CONST_DIV	    	   21
#define ML_CONST_FILL   	   22
#define ML_CONST_SHIFT   	   23

// Ops on source image
#define ML_NOT          	   24
#define ML_NEG          	   25
#define ML_ABS          	   26

/* Density function to be used to translate the histogram.            */
/* Used with ModiHisto()                                              */
#define  ML_UNI                0
#define  ML_EXP                1
#define  ML_RAY                2
#define  ML_HYP_CUB            3
#define  ML_HYP_LOG            4

/* Maximum line length */
#define PR_MAX_LENGTH    1024

/* Function proto types */
#ifdef ANSI 
#ifdef __cplusplus
extern "C" {
#endif
PR_SHORT FTYPE mlProcSetImgSrc   ( PR_SHORT surface_id, PR_SHORT x, 
                                   PR_SHORT y, PR_USHORT xsize, 
                                   PR_USHORT ysize );
PR_VOID  FTYPE mlProcGetImgSrc   ( PR_SHORT  PTYPE *surf_type, 
                                   PR_SHORT  PTYPE *x, 
                                   PR_SHORT  PTYPE *y, 
                                   PR_USHORT PTYPE *xsize, 
                                   PR_USHORT PTYPE *ysize,
                                   PR_UCHAR  PTYPE *pix_size );
PR_SHORT FTYPE mlProcSetImgDest  ( PR_SHORT surface_id, PR_SHORT x, 
                                   PR_SHORT y ); 
PR_VOID  FTYPE mlProcGetImgDest  ( PR_SHORT PTYPE *surf_type, 
                                   PR_SHORT PTYPE *x, 
                                   PR_SHORT PTYPE *y,
                                   PR_UCHAR PTYPE *pix_size ); 
PR_VOID  FTYPE mlProcSetObjCol   ( PR_UCHAR low, PR_UCHAR high );
PR_VOID  FTYPE mlProcGetObjCol   ( PR_UCHAR PTYPE *low, PR_UCHAR PTYPE *high );
PR_VOID  FTYPE mlProcConv3x3     ( PR_UCHAR func_type, PR_CHAR PTYPE *kernel );
PR_VOID  FTYPE mlProcConv3x3Mix  ( PR_UCHAR func_type, PR_CHAR PTYPE *kernel1, 
                                   PR_CHAR PTYPE *kernel2);
PR_VOID  FTYPE mlProcCalcProf    ( PR_ULONG PTYPE *x_profile, 
                                   PR_ULONG PTYPE *y_profile );
PR_VOID  FTYPE mlProcCalcHisto   ( PR_ULONG PTYPE *histo_buf );
PR_VOID  FTYPE mlProcModifHisto  ( PR_UCHAR funct, PR_UCHAR min, 
                                   PR_UCHAR max, PR_ULONG alpha, 
                                   PR_ULONG PTYPE *histobuf, 
                                   PR_UCHAR PTYPE *trans_buf );
PR_VOID  FTYPE mlProcLutMap      ( PR_UCHAR PTYPE *trans_buf );
PR_VOID  FTYPE mlProcErodeImg    ( PR_VOID );
PR_VOID  FTYPE mlProcDilateImg   ( PR_VOID );
PR_VOID  FTYPE mlProcCleanImg    ( PR_VOID );
PR_VOID  FTYPE mlProcArith       ( PR_UCHAR op, PR_LONG constant );
PR_UCHAR FTYPE mlProcLabelImg    ( PR_VOID );
PR_SHORT FTYPE mlProcBufAlloc    ( PR_SHORT surface_type, PR_USHORT xsize, 
                                   PR_USHORT ysize, PR_UCHAR pixsize );
PR_SHORT FTYPE mlProcBufFree     ( PR_SHORT surface_id );
PR_SHORT FTYPE mlProcInit        ( PR_VOID );
PR_SHORT FTYPE mlProcClose       ( PR_VOID );
PR_SHORT FTYPE mlProcBufCpy      ( PR_SHORT  src_id, PR_SHORT  src_x,  
                                   PR_SHORT  src_y,  PR_SHORT  dest_id, 
                                   PR_SHORT  dest_x, PR_SHORT  dest_y, 
                                   PR_USHORT xsize,  PR_USHORT ysize );
PR_SHORT FTYPE mlProcBufWrArea   ( PR_SHORT surface_id, 
                                   PR_SHORT x, PR_SHORT y, 
										     PR_USHORT xsize, PR_SHORT ysize,
                                   PR_VOID PTYPE *buf );
PR_SHORT FTYPE mlProcBufRdArea   ( PR_SHORT surface_id, 
                                   PR_SHORT x, PR_SHORT y, 
										     PR_USHORT xsize, PR_SHORT ysize,
                                   PR_VOID PTYPE *buf );
PR_SHORT FTYPE	mlProcBufCreateId ( PR_VOID PTYPE *bufptr, PR_USHORT xsize, 
											  PR_USHORT ysize, PR_UCHAR pixelsize );
PR_VOID PTYPE * FTYPE mlProcBufGetPtr   ( PR_SHORT bufid );

#ifdef __cplusplus
}
#endif

#else
PR_SHORT FTYPE mlProcSetImgSrc      ();
PR_VOID  FTYPE mlProcGetImgSrc      ();
PR_SHORT FTYPE mlProcSetImgDest     ();
PR_VOID  FTYPE mlProcGetImgDest     ();
PR_VOID  FTYPE mlProcSetObjCol      ();
PR_VOID  FTYPE mlProcGetObjCol      ();
PR_VOID  FTYPE mlProcConv3x3        ();
PR_VOID  FTYPE mlProcConv3x3Mix     ();
PR_VOID  FTYPE mlProcCalcProf       ();
PR_VOID  FTYPE mlProcCalcHisto      ();
PR_VOID  FTYPE mlProcModifHisto     ();
PR_VOID  FTYPE mlProcLutMap         ();
PR_VOID  FTYPE mlProcErodeImg       ();
PR_VOID  FTYPE mlProcDilateImg      ();
PR_VOID  FTYPE mlProcCleanImg       ();
PR_VOID  FTYPE mlProcArith          ();
PR_UCHAR FTYPE mlProcLabelImg       ();
PR_SHORT FTYPE mlProcBufAlloc       ();
PR_SHORT FTYPE mlProcBufFree        ();
PR_SHORT FTYPE mlProcInit           ();
PR_SHORT FTYPE mlProcClose          ();
PR_SHORT FTYPE mlProcBufCpy         ();
PR_SHORT FTYPE mlProcBufWrArea      ();
PR_SHORT FTYPE mlProcBufRdArea      ();
PR_SHORT FTYPE	mlProcBufCreateId    ();
PR_VOID PTYPE * FTYPE mlProcBufGetPtr ();

#endif /* ANSI */						

/* External definitions for predefined kernels */
/* Single convolution */
PR_IMPORT PR_CHAR PTYPE _ML_SH1Ker [ 9 ];
PR_IMPORT PR_CHAR PTYPE _ML_SH2Ker[ 9 ]; 
PR_IMPORT PR_CHAR PTYPE _ML_LP1Ker[ 9 ]; 
PR_IMPORT PR_CHAR PTYPE _ML_LP2Ker[ 9 ]; 
PR_IMPORT PR_CHAR PTYPE _ML_HoriKer[ 9 ];
PR_IMPORT PR_CHAR PTYPE _ML_VertKer[ 9 ];
PR_IMPORT PR_CHAR PTYPE _ML_AverageKer[ 9 ];
PR_IMPORT PR_CHAR PTYPE _ML_SmoothKer[ 9 ];

/* Mix convolution */
PR_IMPORT PR_CHAR PTYPE _ML_PW1Ker[ 9 ];
PR_IMPORT PR_CHAR PTYPE _ML_PW2Ker[ 9 ];
PR_IMPORT PR_CHAR PTYPE _ML_SB1Ker[ 9 ];
PR_IMPORT PR_CHAR PTYPE _ML_SB2Ker[ 9 ];

#endif /* __PR_PROT_H__	*/
