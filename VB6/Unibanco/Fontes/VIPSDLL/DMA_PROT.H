/*
    filename:   dma_prot.h
    Synopsis:   Include file for dma library.
    Date:       06/01/93

COPYRIGHT (c) 1993 Matrox Electronic Systems Ltd.
All Rights Reserved
*/

#ifndef __DMA_PROT_H__
#define __DMA_PROT_H__

//Some types defined here are used with prototypes
#ifndef __EXEC2_H__     //Don't include the standard library includes in
                        //MATROX internal test programs
#include <stdio.h>
#include <stdlib.h>
#endif

// Define ANSI
#ifndef ANSI
#define ANSI
#endif /* ANSI */

// DMA defines
#define ML_DMA_VERIFY 0x00	 // Do not activate any WRITE or READ signals
#define ML_DMA_WRITE  0x01	 // Write to RAM from DMA device
#define ML_DMA_READ   0x02	 // Read from RAM to DMA device

// ****** Function prototype
#ifdef  ANSI
#ifdef __cplusplus
extern "C" {
#endif
short FTYPE mlDmaWrite       ( short surf, short surfx, short surfy,
									    unsigned short xsize, unsigned short ysize );
short FTYPE mlDmaRead        ( short surf, short surfx, short surfy,
									    unsigned short xsize, unsigned short ysize );
short FTYPE mlDmaBufAlloc    ( unsigned short xsize, unsigned short ysize,
                               unsigned char pixelsize );
short FTYPE mlDmaBufFree     ( short bufid );
short FTYPE mlDmaBufSelect   ( short bufid );
short FTYPE mlDmaBufSetPos   ( short x, short y );
short FTYPE mlDmaBufDeselect ( void );
short FTYPE mlDmaInit        ( void );
short FTYPE mlDmaClose       ( void );
short FTYPE mlDmaWait        ( void );
short FTYPE mlDmaWaitLoop    ( unsigned short loop_cnt );
short FTYPE mlDmaBufWrArea   ( short bufid, short x, short y,
									    unsigned short xsize, unsigned short ysize,
									    void PTYPE *buffer );
short FTYPE mlDmaBufRdArea   ( short bufid, short x, short y,
									    unsigned short xsize, unsigned short ysize,
									    void PTYPE *buffer );
void PTYPE * FTYPE mlDmaBufGetPtr   ( short bufid );
void FTYPE mlDmaEnDigSync           ( void );
void FTYPE mlDmaDisDigSync          ( void );
short FTYPE mlDmaBufCreateId( unsigned long physaddr, unsigned short xsize,
                              unsigned short ysize,   unsigned char pixelsize);

// New prototypes that for larger dma buffers
short FTYPE mlDmaWriteTrans     ( short surf, unsigned long surfx, unsigned long surfy,
									       unsigned long xsize, unsigned long ysize );
short FTYPE mlDmaReadTrans      ( short surf, unsigned long surfx, unsigned long surfy,
									       unsigned long xsize, unsigned long ysize);
#ifdef MSC_NT
unsigned long   FTYPE mlDmaPerf( PDMA_PERF_INFO  pDmaPerf ) ;
#endif /* MSC_NT */
short FTYPE mlDmaBufferAlloc    ( unsigned long xsize, unsigned long ysize,
                                  unsigned char pixelsize );
short FTYPE mlDmaBufferSetPos   ( unsigned long x, unsigned long y );
short FTYPE mlDmaBufferWrArea   ( short bufid, unsigned long x, unsigned long y,
									       unsigned long xsize, unsigned long ysize,
									       void PTYPE *buffer );
short FTYPE mlDmaBufferRdArea   ( short bufid, unsigned long x, unsigned long y,
								          unsigned long xsize, unsigned long ysize,
									       void PTYPE *buffer );
short FTYPE mlDmaBufferCreateId ( unsigned long physaddr, unsigned long xsize,
                                  unsigned long ysize,   unsigned char pixelsize);
void PTYPE * FTYPE mlDmaBufferGetPtr ( short bufid );
short FTYPE mlDmaBufferFree     ( short bufid );
short FTYPE mlDmaBufferSelect   ( short bufid );
short FTYPE mlDmaBufferDeselect ( void );

// specific functions
short FTYPE mlDmaBufferLock     (short bufid);
short FTYPE mlDmaBufferUnlock   (short bufid);
unsigned long FTYPE mlDmaBufferGetPhysAdr (short bufid);
short FTYPE mlDmaBufferFreeId   (short bufid);
short FTYPE mlDmaBufferAssignId (unsigned char PTYPE * virtptr,
                                unsigned long physaddr, unsigned long xsize,
                                unsigned long ysize, unsigned char pixelsize);


#ifdef __cplusplus
}
#endif
#else

short FTYPE mlDmaWrite       ( );
short FTYPE mlDmaRead        ( );
short FTYPE mlDmaWriteB      ( );
short FTYPE mlDmaReadB       ( );
short FTYPE mlDmaBufAlloc    ( );
short FTYPE mlDmaBufAlloc2   ( );
short FTYPE mlDmaBufFree     ( );
short FTYPE mlDmaBufSelect   ( );
short FTYPE mlDmaBufSetPos   ( );
short FTYPE mlDmaBufSetPos2  ( );
short FTYPE mlDmaBufDeselect ( );
short FTYPE mlDmaInit        ( );
short FTYPE mlDmaClose       ( );
short FTYPE mlDmaWait        ( );
short FTYPE mlDmaBufWrArea   ( );
short FTYPE mlDmaBufRdArea   ( );
void PTYPE * FTYPE mlDmaBufGetPtr   ( );
void FTYPE mlDmaEnDigSync    ( );
void FTYPE mlDmaDisDigSync   ( );
short FTYPE mlDmaBufCreateId ( );


// New prototypes that for larger dma buffers
short FTYPE mlDmaWriteTrans  ( );
short FTYPE mlDmaReadTrans   ( );
short FTYPE mlDmaBufferAlloc ( );
short FTYPE mlDmaBufferSetPos( );
short FTYPE mlDmaBufferWrArea( );
short FTYPE mlDmaBufferRdArea( );
short FTYPE mlDmaBufferCreateId ( );
short FTYPE mlDmaBufferFree     ( );
short FTYPE mlDmaBufferSelect   ( );
short FTYPE mlDmaBufferDeselect ( );
short FTYPE mlDmaBufferLock  ( );
short FTYPE mlDmaBufferUnlock   ( );
unsigned long FTYPE mlDmaBufferGetPhysAdr ( );
short FTYPE mlDmaBufferFreeId   ( );
short FTYPE mlDmaBufferAssignId ( );



#endif /* ANSI */

#endif /* __DMA_PROT_H__ */




