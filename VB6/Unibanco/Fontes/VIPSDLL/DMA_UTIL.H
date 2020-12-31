/*
    filename:   dma_util.h
    Synopsis:   Include file for dma utilities. The functions
                prototyped in this file should be placed in a tsr for
                windows and dos and in a driver for unix, OS/2 and
					 windows NT operating systems.  
    Date:       06/01/93

COPYRIGHT (c) 1993 Matrox Electronic Systems Ltd.
All Rights Reserved
*/

#ifndef __DMA_UTIL_H__
#define __DMA_UTIL_H__

// Include DMA TSR defines 
#include "tsrfunc.h"

// Define ANSI
#ifndef ANSI
#define ANSI
#endif /* ANSI */

#define PR_BLKSIZE 65536L 

#ifdef BORLAND_32
  #ifdef outp
     #undef outp
     #undef inp
     #undef outportb
     #undef inportb
  #endif
#endif  
     
#ifdef MSC_COMP
   #define INPORTB(port) _inp((unsigned)(port))
   #define OUTPORTB(port, databyte) _outp((unsigned)(port), (int)(databyte))
   #define HUGEBUFALLOC(num, size) _halloc(((unsigned long) (size)), (size_t)(num))
   #define HUGEBUFFREE(h_ptr)     _hfree((void huge *) (h_ptr))
#else
   #ifdef WATCOM
      #define INPORTB(port) inp((unsigned)(port))
      #define OUTPORTB(port, databyte) outp((unsigned)(port), (unsigned char) (databyte))
   #else 
      #define INPORTB(port) inportb((unsigned)(port))
      #define OUTPORTB(port, databyte) outportb((unsigned)(port), (unsigned char) (databyte))
   #endif
   #define HUGEBUFALLOC(num, size) farcalloc((unsigned long) (num), (unsigned long) (size))
   #define HUGEBUFFREE(h_ptr)      farfree( (void far *) (h_ptr))
#endif

/* DPMI Services defines */
#define DPMI_SERVICES                         0x31
#define DPMI_ALLOC_LDT_SELECTOR               0x0
#define DPMI_FREE_LDT_SELECTOR                0x1
#define DPMI_SEGMENT_TO_DESCRIPTOR            0x2
#define DPMI_GET_SELECTOR_INCR                0x3
#define DPMI_GET_SELECTOR_BASE                0x6
#define DPMI_SET_SELECTOR_BASE                0x7
#define DPMI_SET_SELECTOR_LIMIT               0x8
#define DPMI_GET_DESCRIPTOR                   0xb
#define DPMI_CREATE_ALIAS_DESCR 		          0xa
#define DPMI_SIMULATE_REAL_MODE_INT           0x300
#define DPMI_ALLOCATE_DOS_MEMORY              0x100
#define DPMI_FREE_DOS_MEMORY                  0x101
#define DPMI_CALL_REAL_MODE_PROC_WITH_FAR_RET 0x301
#define DPMI_ALLOC_MEM_BLOCK                  0x501
#define DPMI_FREE_MEM_BLOCK                   0x502
#define DPMI_ALLOCATE_LINEAR_MEMORY_BLOCK     0x504
#define DPMI_LOCK_LIN_REGION                  0x600
#define DPMI_UNLOCK_LIN_REGION                0x601
#define DPMI_MAP_PHYSICAL_ADD                 0x800
#define DPMI_FREE_PHYSICAL_ADD                0x801

/* XMS Services defines */
#define MULTIPLEX_SERVICES   0x2F
#define XMS_PRESENT          0x4300
#define XMS_ADDRESS          0x4310

#define XMS_INQUIRE_MEMORY   0x08
#define XMS_ALLOC_MEMORY     0x09
#define XMS_LOCK_MEMORY      0x0C
#define XMS_UNLOCK_MEMORY    0x0D
#define XMS_FREE_MEMORY      0x0A

// Size of selectors 
#define MAXSELSIZE  0x00010000L

// Define for maximum number of dma buffer structures
#define MAXBUF  0x40

// VDS structure 
typedef struct vds_struct {
						unsigned long regionsize;
						unsigned long offset;
						unsigned short segment;
						unsigned short buffid;
						unsigned long  physaddr;
		  } VDS;


// DMA buffer structure
typedef struct dma_bufparms {
					  unsigned char PTYPE *virtptr;    // Virtual pointer 
					  unsigned char       alias;       // Alias buffer id flag
     				           long       x1;			  // Coords
					           long       y1;
								  long       x2;
								  long       y2;
					  unsigned long       xsize;
					  unsigned long       ysize;
					  unsigned char       pixelsize;   // Pixel depth
					  unsigned long       num_sel;     // Number of selectors
					  unsigned long       mem_handle;  // DPMI memory handle
					  unsigned char       acc_type;	  // Read or Write
					  unsigned char       locked;      // Lock flag
					  unsigned long       lin_add;     // linear address
					  VDS                 vds;	     	  // Vds structure;
		  } DMA_BUFPARMS;

// VDS defines
#define VDS_SERVICES                   0x4B    
#define VDS_GET_VERSION                0x8102
#define VDS_LOCK_REGION                0x8103
#define VDS_UNLOCK_REGION              0x8104
#define VDS_SCAT_GATHER_LOCK_REGION    0x8105
#define VDS_SCAT_GATHER_UNLOCK_REGION  0x8106
#define VDS_REQUEST_DMA_BUFFER         0x8107
#define VDS_RELEASE_DMA_BUFFER         0x8108
#define VDS_COPY_TO_DMA_BUFFER         0x8109
#define VDS_COPY_OUT_DMA_BUFFER        0x810A
#define VDS_DISABLE_DMA_TRANSLATION    0x810B
#define VDS_ENABLE_DMA_TRANSLATION     0x810C


/*************************************/
/* Port for DMA controller Register. */
/*************************************/
#define ClearBytePointer(X)       (OUTPORTB(((((X)>>2)&1)?0x00D8:0x000C),0))
#define MasterClear(X)            (OUTPORTB(((((X)>>2)&1)?0x00DA:0x000D),0))
#define ClearMaskRegister(X)      (OUTPORTB(((((X)>>2)&1)?0x00DC:0x000E),0))

#define InterruptOnChannel(X)     (inportb(0x040A)&(1<<(X)))

#define BaseWordCountLow          (0x0007)
#define BaseWordCountHigh         (0x0407)

#define BaseAddress8237           (0x0006)
#define BaseAddressLow            (0x0082)
#define BaseAddressHigh           (0x0482)

/* ************************************************************************ */
/*                                                                          */
/* DMA Extended Mode Register                             rev 3.12, pg. 254 */
/*                                                                          */
/* Used to program the DMA data size and timing mode.                       */
/*                                                                          */
/*                                                                          */
/* ************************************************************************ */

typedef struct DMAExtendedMode {
  unsigned DMAChannelSelect       : 2;
  unsigned AddressingMode         : 2;
  unsigned DMACycleTimingMode     : 2;
  unsigned TCSignalDirection      : 1;
  unsigned StopRegister           : 1;
} DMAExtendedMode_T;   

typedef struct DMAChainingMode {
  unsigned DMAChannelSelect       : 2;
  unsigned ChainingModeControl    : 2;
  unsigned CompletionSignal       : 1;
} DMAChainingMode_T;

typedef struct DMAMode {
  unsigned DMAChannelSelect       : 2;
  unsigned DataTransferType       : 2;
  unsigned AutoInitialization     : 1;
  unsigned AddressIncrementSelect : 1;
  unsigned DMAChannelModeSelect   : 2;
} DMAMode_T;



#define DmaExtendedMode(X)        ((((X)>>2)&1)?0x04D6:0x040B)
#define  ADDR(X)                  (((X)&3)<<2)
#define   B8                      (0)
#define   W16                     (1)
#define   B32                     (2)
#define   B16                     (3)
#define  CYCLE(X)                 (((X)&3)<<4)
#define   TYPE_ISA                (0)         
#define   TYPE_A                  (1)
#define   TYPE_B                  (2)
#define   TYPE_C                  (3)
#define   BURST                   (TYPE_C)
#define  T_C(X)                   (((X)&1)<<6)
#define   OUTPUT                  (0)
#define   INPUT                   (1)
#define  STOP(X)                  (((X)&1)<<7)

/* ************************************************************************ */
/*                                                                          */
/* Chaining Mode Register                                 rev 3.12, pg. 256 */
/*                                                                          */
/* Used to control buffer chaining.                                         */
/*                                                                          */
/*                                                                          */
/* ************************************************************************ */

#define ChainingMode(X)           ((((X)>>2)&1)?0x04D4:0x040A)
#define  CHAIN(X)                 (((X)&3)<<2)
#define   CHAIN_STOP              (0)
#define   CHAIN_LOAD              (1)
#define   CHAIN_UPDATE            (3)
#define  NOTIFY(X)                (((X)&1)<<4)
#define   SW                      (0)         
#define   HW                      (1)

#define  ChainingModeStatus(X)     (0x04D4)
#define  ChannelInterruptStatus(X) (0x040A)
#define  ChainBufferExpiration(X)  (0x040C)

#define WriteSingleMaskBit(X)     ((((X)>>2)&1)?0x00D4:0x000A)
#define  DMA(X)                   (((X)&1)<<2)
#define   ENABLED                 (0)
#define   DISABLED                (1)

#define MaskStatus(X)             (((inportb((((X)>>2)&1)?0x00DE:0x000F) )>>((X)&3))&1)

#define OperationControlWord      (0x00A0)                    
#define  DMA(X)                   (((X)&1)<<2)
#define   ENABLED                 (0)
#define   DISABLED                (1)


#define ModeRegister(X)           ((((X)>>2)&1)?0x00D6:0x000B)
#define  TYPE(X)                  (((X)&3)<<2)
#define   MVERIFY                  (0)
#define   MWRITE                   (1)
#define   MREAD                    (2)
#define  AUTO(X)                  (((X)&1)<<4)
#define   OFF                     (0)
#define   ON                      (1)
#define  DIR(X)       (((X)&1)<<5)
#define   INC          0
#define   DEC          1
#define  MODE(X)      (((X)&3)<<6)
#define   DEMAND       0
#define   SINGLE       1
#define   BLOCKMODE    2
#define   CASCADE      3

#define StatusRegister(X)           ((((X)>>2)&1)?0x00D0:0x0008)
#define  TC_STATUS(X)               (1<<((X)&0x03))

/***********************/
/* DMA Register Values */
/***********************/
#define CHANNEL(X)   ((X)&3)

   /****************************************************/
   /* Function number that we can find in register AH  */
   /* for execute when we have a Magic interruption.   */
   /****************************************************/
   #define UNLOAD_TSR        0x00
   #define BOARD_INFO        0x01
   #define GET_VERSION       0x02
   #define BOARD_RESET       0x03
   #define CAMERA_LOAD       0x04
   #define CFG_DONE          0x05
   #define REG_MEM           0x06
   #define GOUT_PORT         0x07
   #define ASSIGN_PROCESSID  0x08
   #define UPDATE_FINISHED   0x09
   #define FREE_PROCESSID    0x0A
   #define WAIT_FINISHED     0x0B
   #define CURRENT_PROCESSID 0x0C

// Function prototypes 
short InitSelectors( unsigned long selectorbase, unsigned long lin_add, 
                     unsigned long size );
#ifdef DMA_STANDALONE
short FTYPE Dma_SegToDsc  ( void PTYPE * PTYPE * real_ptr );
short FTYPE Dma_CallTsr   ( union REGS PTYPE * regs, 
                            struct SREGS PTYPE * sregs);
#endif /* DMA_STANDALONE */

#endif /* __DMA_UTIL_H__ */
