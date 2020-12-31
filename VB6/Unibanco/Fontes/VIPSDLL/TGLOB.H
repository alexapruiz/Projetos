/*    filename:   tglob.h

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved

*/

#ifndef __TGLOB_H__
#define __TGLOB_H__

#ifndef __TSERIES_H__
#include <tseries.h>
#endif /* __TSERIES_H__ */

#include <tcamera.h>

#define  MABOARD  struct  MagicBoard
#define  THRA     struct  hra
#define  NBMAXERR    15
#define  MADISP   struct  MagicTsrDisp

typedef  struct t_glob {
   // number of boards present in system
   unsigned char        NbBoard;
   // current board
   short                SelBoard;

   unsigned short TsrVector;

   // pointer to TSR data area
   struct MagicTsrBoard (MFAR_1M *MaTsr)[T_N_OF_BOARDS];
   struct MagicAdjAtt   (MFAR_1M *MaTsrAtt)[T_N_OF_BOARDS];

   // pointer to current TCAMERA structure;
   TCAMERA        MFAR* SelCamera;
   // An internal test feature...
   TENHMAGIC      MFAR* EnhMagicReg;

   // pointers to HW
   MABOARD        board[T_N_OF_BOARDS];
   SHADOW         shade[T_N_OF_BOARDS];
   SHADOW         MFAR* shadow;
   volatile HRA            MFAR_1M* hra;
   volatile unsigned char  MFAR_1M* mlb1;
   volatile unsigned char  MFAR_1M* mlb2;
   volatile unsigned char  MHUGE* HighMapSel;

   // current DMA channel number;
   unsigned short DmaChannel;

   unsigned short BoardIsMapped;

   // Pointer to variables defined in the TSR
   volatile unsigned long MFAR_1M * dmafinished;
   volatile unsigned long MFAR_1M * grabfinished;
   volatile unsigned long MFAR_1M * vsyncfinished;
   volatile unsigned long MFAR_1M * realgrabdone;
   volatile unsigned long MFAR_1M * realdmadone;
   volatile unsigned long MFAR_1M * realvsyncdone;
   volatile unsigned long MFAR_1M * realdispvsyncdone;
   volatile unsigned long MFAR_1M * realhsyncdone;
   volatile unsigned long MFAR_1M * virtgrabdone;
   volatile unsigned long MFAR_1M * virtdmadone;
   volatile unsigned long MFAR_1M * virtvsyncdone;
   volatile unsigned long MFAR_1M * virtdispvsyncdone;
   volatile unsigned long MFAR_1M * virthsyncdone;

   // Pointers to the starting addresses of arrays
   // defined in the TSR

   volatile unsigned long MFAR_1M * str_dmafinished;
   volatile unsigned long MFAR_1M * str_grabfinished;
   volatile unsigned long MFAR_1M * str_vsyncfinished;
   volatile unsigned long MFAR_1M * str_realgrabdone;
   volatile unsigned long MFAR_1M * str_realvsyncdone;
   volatile unsigned long MFAR_1M * str_realhsyncdone;
   volatile unsigned long MFAR_1M * str_realdmadone;
   volatile unsigned long MFAR_1M * str_realdispvsyncdone;
   volatile unsigned long MFAR_1M * str_virtgrabdone;
   volatile unsigned long MFAR_1M * str_virtdmadone;
   volatile unsigned long MFAR_1M * str_virthsyncdone;
   volatile unsigned long MFAR_1M * str_virtvsyncdone;
   volatile unsigned long MFAR_1M * str_virtdispvsyncdone;


   // pointer to vertical display sync information
   volatile MADISP        MFAR_1M * DispInfoStr;
   volatile MADISP        MFAR_1M * DispInfo;

   // flag for the vertical sync interrupt status
   unsigned short dispVsync[T_N_OF_BOARDS];

   unsigned short camLoad[T_N_OF_BOARDS];

   // flags for grab
   unsigned short initFlag[T_N_OF_BOARDS];
   unsigned short grabVal[T_N_OF_BOARDS];
   unsigned short grabType[T_N_OF_BOARDS];
   unsigned short grabStatus[T_N_OF_BOARDS];
   long     pedestal[T_N_OF_BOARDS];
   } T_GLOB;



   //
   // stuff for Windows NT...
   //
typedef struct _T_DEVICE {
   // handlers for Windows NT driver
   void           *hMagicDevice[T_N_OF_BOARDS];

   // flag for multitasking operation under Windows NT
   unsigned short DeviceAllocated[T_N_OF_BOARDS];
   unsigned short FirstDevice;

   // pointer to the info struct returned by the Windows NT driver
   void           *pDeviceInfo;

   // handle to the current process
   void           *hProcess;

   // id of the current process
   unsigned long   ProcessId;

   // function mlAllocateDevice() has been called
   unsigned short mlAllocateDeviceCalled;
   } T_DEVICE;


typedef struct t_user {
   // host access and memory organization
   char           memOrg;
   unsigned int   memType;
   unsigned int   pixSize;
   unsigned int   lineLength;
   long           pitch;
   unsigned int   busPixSize;

   // display setup
   char MFAR_1M * S3DualScreenRegBase;
   unsigned short dispXres;
   unsigned short dispYres;

   // digitizer setup
   unsigned short iterations;
   unsigned short condition;
   unsigned short field;
   unsigned short grabtrig;
   unsigned short surface;
   unsigned short posX;
   unsigned short posY;
   unsigned short sizeX;
   unsigned short sizeY;
   unsigned short subX;
   unsigned short subY;
   unsigned short psg_vsbrst;
   unsigned short psg_vebrst;
   unsigned short psg_hsblnk;
   unsigned short psg_heblnk;
   unsigned short windowX;
   unsigned short windowY;
   unsigned short digintlc;
   unsigned short sizeX_in_dcf;
   unsigned short GrabInProgress;
   unsigned short DmaInProgress;
   unsigned short useBT812;
   unsigned short bt812_hdelay;
   unsigned short bt812_vdelay;

   // version of the TSR.
   unsigned short tsrversion;
   unsigned short specialMode;

   short dispMode;
   short xStart;
   short yStart;
   } T_USER;

#ifdef BORLAND_32
typedef struct { unsigned long DeviceRegAddr[T_N_OF_BOARDS];
                 unsigned long DeviceHighMap[T_N_OF_BOARDS];
                 unsigned long grabfinished;
                 unsigned long grabdone;
                 unsigned long dmafinished;
                 unsigned long dmadone;
                 unsigned long vsyncfinished;
                 unsigned long vsyncdone;
                 unsigned long hsyncdone;
                 unsigned long dispvsyncdone;
                 unsigned long S3SegRegMap[T_N_OF_BOARDS];
                 unsigned long MaTsr;
                 unsigned long MaTsrAtt;
                 unsigned long DispInfoStr;
               } HANDLESTRUCT;

#endif /* BORLAND_32  */
#endif /* __TGLOB_H__ */



