/*++

Copyright (c) 1994  Matrox Electronic System

Author:

    Daniel Pettigrew


Module Name:

    magic.h

Abstract:


Environment:

    kernel & User mode

Notes:


Revision History:
    94/08/01   dpet        create the file

--*/
#ifndef __MAGICDRV_H__
#define __MAGICDRV_H__


//
// Define the various device type values.  Note that values used by Microsoft
// Corporation are in the range 0-32767, and 32768-65535 are reserved for use
// by customers.
//

#define FILE_DEVICE_MAGIC  45634


//
// Macro definition for defining IOCTL and FSCTL function control codes.  Note
// that function codes 0-2047 are reserved for Microsoft Corporation, and
// 2048-4095 are reserved for customers.
//

#define MAGIC_IOCTL_INDEX  3768


//
// Define our own private IOCTL
//

#define IOCTL_MAGIC_MAP_USER_PHYSICAL_MEMORY     CTL_CODE(FILE_DEVICE_MAGIC , \
                                                           MAGIC_IOCTL_INDEX, \
                                                           METHOD_BUFFERED,   \
                                                           FILE_ANY_ACCESS)

#define IOCTL_MAGIC_UNMAP_USER_PHYSICAL_MEMORY   CTL_CODE(FILE_DEVICE_MAGIC,  \
                                                          MAGIC_IOCTL_INDEX+1,\
                                                          METHOD_BUFFERED,    \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_READ_PORT_UCHAR              CTL_CODE(FILE_DEVICE_MAGIC , \
                                                          MAGIC_IOCTL_INDEX+2,\
                                                          METHOD_BUFFERED,    \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_WRITE_PORT_UCHAR             CTL_CODE(FILE_DEVICE_MAGIC,  \
                                                          MAGIC_IOCTL_INDEX+3,\
                                                          METHOD_BUFFERED,    \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_OBTAIN_DEVICE_INFO_ADDRESS   CTL_CODE(FILE_DEVICE_MAGIC,  \
                                                          MAGIC_IOCTL_INDEX+4,\
                                                          METHOD_BUFFERED,    \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_RELEASE_DEVICE_INFO_ADDRESS  CTL_CODE(FILE_DEVICE_MAGIC,  \
                                                          MAGIC_IOCTL_INDEX+5,\
                                                          METHOD_BUFFERED,    \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_FPGA_CFG_DONE                CTL_CODE(FILE_DEVICE_MAGIC,  \
                                                          MAGIC_IOCTL_INDEX+6,\
                                                          METHOD_BUFFERED,    \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_ACCESS_GOUTPORT              CTL_CODE(FILE_DEVICE_MAGIC,  \
                                                          MAGIC_IOCTL_INDEX+7,\
                                                          METHOD_BUFFERED,    \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_DISPVSYNC_INTERRUPT          CTL_CODE(FILE_DEVICE_MAGIC,  \
                                                          MAGIC_IOCTL_INDEX+8,\
                                                          METHOD_BUFFERED,    \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_INTERRUPT_WAIT               CTL_CODE(FILE_DEVICE_MAGIC,  \
                                                          MAGIC_IOCTL_INDEX+9,\
                                                          METHOD_BUFFERED,    \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_OBTAIN_DEVICE_INFO_DATA      CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+10,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_DMA_WRITE_TRANSFER           CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+11,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_DMA_READ_TRANSFER            CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+12,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_DMA_INFO_PARAMETERS          CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+13,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_DMA_BUFFER_ALLOC             CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+14,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_DMA_BUFFER_FREE              CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+15,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_GET_DMA_PERF                 CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+16,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_PROVIDE_PROCESS_INFO         CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+17,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_DMA_BUFFER_SELECT            CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+18,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_DMA_BUFFER_SET_POS           CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+19,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_DMA_BUFFER_GET_PTR           CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+20,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)

#define IOCTL_MAGIC_DMA_BUFFER_GET_INFO          CTL_CODE(FILE_DEVICE_MAGIC,   \
                                                          MAGIC_IOCTL_INDEX+21,\
                                                          METHOD_BUFFERED,     \
                                                          FILE_ANY_ACCESS)


//
//
// Our user mode app will pass an initialized structure like this
//     down to the kernel mode driver
//
//------------------------------------------------------------------------------
typedef struct
{
   ULONG            InterfaceType; // Isa, Eisa, etc....
   ULONG            BusNumber;     // Bus number
   LARGE_INTEGER    BusAddress;    // Bus-relative address
   ULONG            AddressSpace;  // 0 is memory, 1 is I/O
   ULONG            Length;        // Length of section to map
   ULONG            device;        // type of MAGIC.
   PVOID            pMem;          // ptr to the region to unmap

} PHYSICAL_MEMORY_INFO, *PPHYSICAL_MEMORY_INFO;


//------------------------------------------------------------------------------
typedef struct _PORT_ACCESS_INFO
{
   USHORT      port;
   UCHAR       value;
}  PORT_ACCESS_INFO, *PPORT_ACCESS_INFO;


//------------------------------------------------------------------------------
#define EVENT_MAX    7

typedef enum _MAGIC_EVENT_TYPE
{
   grab,
   vsync,
   dma,
   dispvsync,
   hsync,
   userin,
   none

}  MAGIC_EVENT_TYPE;


//------------------------------------------------------------------------------
typedef struct _INTERRUPT_WAIT_INFO
{

   MAGIC_EVENT_TYPE     eEventType;
   BOOLEAN              bResetEventWhenEnter;
   MAGIC_EVENT_TYPE     eResetInterruptCountAtExit;
   ULONG                NumberOfEventDesired;
   ULONG                NumberOfInterruptCountDesired;
   ULONG                NumberOfInterruptCountDone;
   ULONG                ValueToPutInInterruptCountAtExit;
   ULONG                time1;
   ULONG                time2;
   ULONG                time3;

} INTERRUPT_WAIT_INFO, *PINTERRUPT_WAIT_INFO;


//------------------------------------------------------------------------------
// Structure containing Interrupt Control variables

typedef struct _INTERRUPT_INFO
{

   ULONG                count[EVENT_MAX];
   PVOID                InterruptObject;
   ULONG                time2;

} INTERRUPT_INFO, *PINTERRUPT_INFO;


//------------------------------------------------------------------------------
typedef struct _MAGIC_REGISTERS
{

   PVOID                base;
   PUSHORT              irqmask;
   PUSHORT              irqclear;
   PUSHORT              irqstatus;
   PUSHORT              srtstarth;
   PUSHORT              srtstartl;
   PUSHORT              srtstartm;
   PUSHORT              vdoctrl;
   PUSHORT              vdoctrlr;
   PUSHORT              dmaxstart;
   PUSHORT              dmaxstop;
   PUSHORT              dmaystart;
   PUSHORT              dmactrl;

} MAGIC_REGISTERS, *PMAGIC_REGISTERS;


//------------------------------------------------------------------------------
typedef struct _DBL
{

   PVOID                      start;
   ULONG                      size;
   PVOID                      next;
   PVOID                      previous;
   SHORT                      BufId;
   HANDLE                     hProcess;
   ULONG                      ProcessId;
   ULONG                      xsize;
   ULONG                      ysize;
   ULONG                      xpos;
   ULONG                      ypos;
   UCHAR                      pixelsize;
   PVOID                      mapAddress;

}  DBL, *PDBL;


//------------------------------------------------------------------------------
typedef struct _DMA_MEMORY_POOL
{

   ULONG                BufferSize;
   LARGE_INTEGER        PhysAddressDmaPool;
   PVOID                pDmaPool;
   PDBL                 pFreeDBL;
   PDBL                 pUsedDBL;

} DMA_MEMORY_POOL, *PDMA_MEMORY_POOL;


//------------------------------------------------------------------------------
typedef struct _DEVICE_INFO
{

   ULONG                   TotBdNb;
   PVOID                   DeviceObject[BOARD_MAX];
   MAGICTSRBOARD           MagicBd[BOARD_MAX];
   MAGICADJATT             MagicBdAtt[BOARD_MAX];
   MAGICDISPINFO           MagicDispInfo[BOARD_MAX];
   INTERRUPT_INFO          MagicIntInfo[BOARD_MAX];
   MAGIC_REGISTERS         MagicReg[BOARD_MAX];
   DMA_MEMORY_POOL         DmaPool;
   PVOID                   pDblSemaphore;

} DEVICE_INFO, *PDEVICE_INFO;


//------------------------------------------------------------------------------
typedef struct _CFGDONE_INFO
{

   ULONG                device[BOARD_MAX];
   ULONG                globalStatus;

} CFGDONE_INFO, *PCFGDONE_INFO;


//------------------------------------------------------------------------------
typedef struct _GOUTPORT_INFO
{

   ULONG                function;
   ULONG                state;

} GOUTPORT_INFO, *PGOUTPORT_INFO;


//------------------------------------------------------------------------------
typedef struct _DISPVSYNC_INFO
{

   USHORT               value;

} DISPVSYNC_INFO, *PDISPVSYNC_INFO;


//------------------------------------------------------------------------------
typedef struct _DEVICE_MEMORY_INFO
{

   PVOID                pMem;

} DEVICE_MEMORY_INFO, *PDEVICE_MEMORY_INFO;


//------------------------------------------------------------------------------
typedef struct _DMA_INFO_PARAMETERS
{

   USHORT               surf;
   ULONG                surfx;
   ULONG                surfy;
   ULONG                xsize;
   ULONG                ysize;
   USHORT               dmactrl;
   USHORT               lineLength;
   USHORT               FrameBufferAccessType;

} DMA_INFO_PARAMETERS, *PDMA_INFO_PARAMETERS;


//------------------------------------------------------------------------------
typedef struct _DMA_BUFFER_ALLOC_INFO
{

   SHORT                BufId;
   ULONG                size;
   ULONG                xsize;
   ULONG                ysize;
   UCHAR                pixelsize;
   HANDLE               hProcess;
   ULONG                ProcessId;

} DMA_BUFFER_ALLOC_INFO, *PDMA_BUFFER_ALLOC_INFO;


//------------------------------------------------------------------------------
typedef struct _DMA_BUFFER_FREE_INFO
{

   SHORT                BufId;
   HANDLE               hProcess;
   ULONG                ProcessId;
   BOOLEAN              FreeAllWith_ProcessId;

} DMA_BUFFER_FREE_INFO, *PDMA_BUFFER_FREE_INFO;


//------------------------------------------------------------------------------
typedef struct _DMA_BUFFER_SELECT_INFO
{

   SHORT                BufId;
   HANDLE               hProcess;
   ULONG                ProcessId;

} DMA_BUFFER_SELECT_INFO, *PDMA_BUFFER_SELECT_INFO;


//------------------------------------------------------------------------------
typedef struct _DMA_BUFFER_SET_POS_INFO
{

   ULONG                xpos;
   ULONG                ypos;

} DMA_BUFFER_SET_POS_INFO, *PDMA_BUFFER_SET_POS_INFO;


//------------------------------------------------------------------------------
typedef struct _DMA_BUFFER_GET_PTR_INFO
{

   SHORT                BufId;
   PVOID                MapViewPtr;
   ULONG                ProcessId;

} DMA_BUFFER_GET_PTR_INFO, *PDMA_BUFFER_GET_PTR_INFO;


//------------------------------------------------------------------------------
typedef struct _DMA_BUFFER_GET_INFO
{

   PVOID                      start;
   ULONG                      size;
   SHORT                      BufId;
   HANDLE                     hProcess;
   ULONG                      ProcessId;
   ULONG                      xsize;
   ULONG                      ysize;
   UCHAR                      pixelsize;
   PVOID                      mapAddress;

} DMA_BUFFER_GET_INFO, *PDMA_BUFFER_GET_INFO;


//------------------------------------------------------------------------------
typedef struct _PROVIDE_PROCESS_INFO
{

   HANDLE               hProcess;
   ULONG                ProcessId;

} PROVIDE_PROCESS_INFO, *PPROVIDE_PROCESS_INFO;



//------------------------------------------------------------------------------
typedef struct _DMA_PERF_INFO
{

   ULONG                time0;      // in dispatch routine
   ULONG                time1;      // in StartIo routine
   ULONG                time2;      // in AdapterControl routine
   ULONG                time3;      // before IoMapTransfer routine
   ULONG                time4;      // before write to DMA registers
   ULONG                time5;      // in DpcForIsr routine
   ULONG                DmaCounter; // value from HalReadDmaCounter

} DMA_PERF_INFO, *PDMA_PERF_INFO;


#endif
