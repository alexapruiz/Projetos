/******************************************************************************/
/*
*
*    Filename:  PULSAR.H
*    Owner   :  Matrox Imaging dept.
*    Rev     :  $Revision:   5.0  $
*    Content :  This file contains the defines that are needed by the user
*               to use the MIL library with the PULSAR.
*
*    COPYRIGHT (c) Matrox Electronic Systems Ltd.
*    All Rights Reserved
*/
/******************************************************************************/

#ifndef __MIL_H
#include <mil.h>
#endif

#ifndef __PULSAR_H
#define __PULSAR_H

/******************************************************************************/
/*                    PULSAR DRIVER RELEASE (5.00)                            */
/*                                                                            */
/*    These are the specifics or not yet released defines for the PULSAR      */
/******************************************************************************/
#define M_IRQ_GLOBAL_OBJECT           125L

#define M_RGB888_ATIMACH64     0xE5E5E500L /* Internal use only */
#define M_RGB888_NORMALVGA     0x39393900L /* Internal use only */

/* User Hook identification*/
/******************************************************************************/
#define ISR_START_OF_FIELD_MASK      0x000000FFL
#define ISR_START_OF_FIELD_BIT       0x00000001L
#define ISR_START_OF_FIELD_ODD_BIT   0x00000002L
#define ISR_START_OF_FIELD_EVEN_BIT  0x00000004L
#define ISR_START_OF_FRAME_BIT       0x00000008L
#define ISR_START_OF_GRAB_BIT        0x00000010L
#define ISR_START_OF_GRAB_FRAME_BIT  0x00000020L

#define ISR_END_OF_FIELD_MASK        0x0000FF00L
#define ISR_END_OF_FIELD_EVEN_BIT    0x00000100L
#define ISR_END_OF_FIELD_ODD_BIT     0x00000200L
#define ISR_END_OF_FIELD_BIT         0x00000400L
#define ISR_END_OF_GRAB_FRAME_BIT    0x00000800L

#define ISR_END_OF_GRAB_MASK        (0x00FF0000L | ISR_END_OF_GRAB_FRAME_BIT) // special case for linescan no sync
#define ISR_END_OF_GRAB_BIT          0x00010000L
                                        

/* PCI device information                                                     */
/******************************************************************************/
#define M_PCI_VENDOR_ID             0x00 // (16 lsb)
#define M_PCI_DEVICE_ID             0x00 // (16 msb)
#define M_PCI_COMMAND               0x01 // (16 lsb)
#define M_PCI_STATUS                0x01 // (16 msb)
#define M_PCI_REVISION_ID           0x02 // (byte 0)
#define M_PCI_CLASS_CODE            0x02 // (byte 1,2,3)
#define M_PCI_LATENCY_TIMER         0x03 // (byte 1)
#define M_PCI_BASE_ADRS0            0x04 // 
#define M_PCI_BASE_ADRS1            0x05 // 
#define M_PCI_INT_LINE              0x0F // (byte 0)
#define M_PCI_INT_PIN               0x0F // (byte 1)

// check Neptune          chipset  DevID = 0x04a3L VendID = 0x8086L
// check Triton III 430VX chipset  DevID = 0x7030L VendID = 0x8086L
// check Triton II  430HX chipset  DevID = 0x1250L VendID = 0x8086L
// check Triton I   430FX chipset  DevID = 0x122dL VendID = 0x8086L
// check Orion      450KX chipset  DevID = 0x84c4L VendID = 0x8086L
// check ppro       440FX chipset  DevID = 0x1237L VendID = 0x8086L
// MGA STORM  DevID = 0x0519L  VendID = 0x102BL 
// SAA7116    DevID = 0x1223L  VendID = 0x8086L 
// DEC BRIDGE DevID = 0x0001L  VendID = 0x1011L 
// IBM BRIDGE DevID = 0x0022L  VendID = 0x1014L
                                                   

/************************************************************************/
/* MIL OPTIONAL COMPILE DEFINES prefixes                                */
/************************************************************************/
#if M_MIL_USE_DOS
   #define PULSARFARTYPE far              
#endif

#if M_MIL_USE_WINDOWS
   #define PULSARFARTYPE far              
#endif

#if M_MIL_USE_DOS_32
   #ifdef __WATCOMC__
      #define PULSARFARTYPE far              
   #else
      #define PULSARFARTYPE 
   #endif
#endif

#if M_MIL_USE_NT_WINDOWS
   #define PULSARFARTYPE                      
#endif


#ifdef __cplusplus
extern "C"
{
#endif
/*==============================================================================*/
/*                                                                              */
/* PULSAR Lib prototypes                                                        */
/*                                                                              */
/*==============================================================================*/
typedef MFTYPE32 long (MFTYPE MPTYPE *PAPPHOOKFCTPTR)(long FHookType, void MPTYPE *FExpansionFlagPtr);
                                      
MFTYPE32 long MFTYPE PappAlloc(         long InitFlag, 
                                        MIL_ID MPTYPE *ApplicationID,
                                        void MPTYPE *ThreadContextPtr);
                                     
MFTYPE32 void MFTYPE PappControl(       long ControlType, 
                                        long ControlFlag,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PappInquire(       long InquireType, 
                                        long MPTYPE *UserVarPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PappGetError(      long ErrorType, 
                                        void MPTYPE *UserPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PappFree(          MIL_ID ApplicationID,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PsysAlloc(         long SystemType, 
                                        long Number, 
                                        long InitFlag, 
                                        MIL_ID MPTYPE *SystemID,
                                        void MPTYPE *ThreadContextPtr);
                                     
MFTYPE32 long MFTYPE PsysInquire(       MIL_ID SystemID,
                                        long InquireType, 
                                        long MPTYPE *UserVarPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PsysControl(       MIL_ID SystemID, 
                                        long ControlType, 
                                        long ControlFlag,
                                        void MPTYPE *ThreadContextPtr);
                                     
MFTYPE32 void MFTYPE PsysConfigAccess(  MIL_ID SystemID,
                                        long VendorId,      
                                        long DeviceId,      
                                        long DeviceNum,     
                                        long OperationFlag, 
                                        long OperationType, 
                                        long Offset,        
                                        long Size,
                                        void *UserArrayPtr,          
                                        void   MPTYPE *ThreadContext);

MFTYPE32 void MFTYPE PsysFree(          MIL_ID SystemID,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PbufAllocColor(    MIL_ID SystemID, 
                                        long   SizeBand,
                                        long   SizeX,
                                        long   SizeY,
                                        long   Type,
                                        long   Attribute,
                                        MIL_ID MPTYPE *IdVarPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PbufAlloc2d(       MIL_ID SystemID, 
                                        long   SizeX,
                                        long   SizeY,
                                        long   Type,
                                        long   Attribute,
                                        MIL_ID MPTYPE *IdVarPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PbufCreateColor(   MIL_ID SystemID, 
                                        long   SizeBand,
                                        long   SizeX,
                                        long   SizeY,
                                        long   Type,
                                        long   Attribute,
                                        long   ControlFlag,  
                                        long   Pitch,        
                                        void   *UserArrayPtr,
                                        MIL_ID MPTYPE *IdVarPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PbufCreate2d(      MIL_ID SystemID, 
                                        long   SizeX,
                                        long   SizeY,
                                        long   Type,
                                        long   Attribute,
                                        void   *UserArrayPtr,
                                        MIL_ID MPTYPE *IdVarPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PbufChildColor(    MIL_ID BufferID, 
                                        long Band, 
                                        MIL_ID MPTYPE *IdVarPtr,
                                        void MPTYPE *ThreadContextPtr);
                                     
MFTYPE32 long MFTYPE PbufChild2d(       MIL_ID BufferID, 
                                        long OffsetX, 
                                        long OffsetY, 
                                        long SizeX, 
                                        long SizeY,
                                        MIL_ID MPTYPE *IdVarPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PbufPut(           MIL_ID BufferID, 
                                        void MPTYPE *UserArrayPtr,
                                        void MPTYPE *ThreadContextPtr);
                                     
MFTYPE32 void MFTYPE PbufPut1d(         MIL_ID BufferID, 
                                        long OffsetX, 
                                        long SizeX, 
                                        void MPTYPE *UserArrayPtr,
                                        void MPTYPE *ThreadContextPtr);
                                     
MFTYPE32 void MFTYPE PbufPut2d(         MIL_ID BufferID, 
                                        long OffsetX, 
                                        long OffsetY, 
                                        long SizeX, 
                                        long SizeY, 
                                        void MPTYPE *UserArrayPtr,
                                        void MPTYPE *ThreadContextPtr);
                                     
MFTYPE32 void MFTYPE PbufPutColor(      MIL_ID BufferID, 
                                        long Format, 
                                        long Band, 
                                        void MPTYPE *UserArrayPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PbufGet(           MIL_ID BufferID, 
                                        void MPTYPE *UserArrayPtr,
                                        void MPTYPE *ThreadContextPtr);
                                     
MFTYPE32 void MFTYPE PbufGet1d(         MIL_ID BufferID, 
                                        long OffsetX, 
                                        long SizeX, 
                                        void MPTYPE *UserArrayPtr,
                                        void MPTYPE *ThreadContextPtr);
                                     
MFTYPE32 void MFTYPE PbufGet2d(         MIL_ID BufferID, 
                                        long OffsetX, 
                                        long OffsetY, 
                                        long SizeX, 
                                        long SizeY, 
                                        void MPTYPE *UserArrayPtr,
                                        void MPTYPE *ThreadContextPtr);
                                     
MFTYPE32 void MFTYPE PbufGetColor(      MIL_ID BufferID, 
                                        long Format, 
                                        long Band, 
                                        void MPTYPE *UserArrayPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PbufInquire(       MIL_ID BufferID,
                                        long InquireType, 
                                        long MPTYPE *UserVarPtr,
                                        void MPTYPE *ThreadContextPtr);
                                        
MFTYPE32 void MFTYPE PbufFree(          MIL_ID BufferID,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PdigAlloc(         MIL_ID SystemID, 
                                        long DeviceNum, 
                                        char MPTYPE *DataFormat, 
                                        long InitFlag, 
                                        MIL_ID MPTYPE *IdVarPtr,
                                        void MPTYPE *ThreadContextPtr);
                                        
MFTYPE32 void MFTYPE PdigFree(          MIL_ID DigitizerID,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PdigGrabContinuous(MIL_ID DigitizerID, 
                                        MIL_ID BufferID,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PdigGrab(          MIL_ID DigitizerID, 
                                        MIL_ID BufferID,
                                        void MPTYPE *ThreadContextPtr);
                                        
MFTYPE32 void MFTYPE PdigHalt(          MIL_ID DigitizerID,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PdigGrabWait(      MIL_ID DigitizerID,
                                        long Flag,
                                        void MPTYPE *ThreadContextPtr);
                                        
MFTYPE32 void MFTYPE PdigControl(       MIL_ID DigitizerID,
                                        long ControlType, 
                                        double ControlValue,
                                        void MPTYPE *ThreadContextPtr);
                                        
MFTYPE32 long MFTYPE PdigInquire(       MIL_ID DigitizerID,
                                        long InquireType, 
                                        void MPTYPE *UserVarPtr,
                                        void MPTYPE *ThreadContextPtr);
                                        
MFTYPE32 void MFTYPE PdigReference(     MIL_ID DigitizerID,
                                        long ReferenceType, 
                                        long ReferenceLevel,
                                        void MPTYPE *ThreadContextPtr);
                                        
MFTYPE32 void MFTYPE PdigChannel(       MIL_ID DigitizerID,
                                        long Channel,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PdigSetLut(        MIL_ID DigitizerID, 
                                        long LutType, 
                                        long Start, 
                                        long NbElem, 
                                        void MPTYPE *LutData, 
                                        long SizeBit, 
                                        long LutResetType,
                                        void MPTYPE *ThreadContextPtr);


MFTYPE32 MDIGHOOKFCTPTR MFTYPE PdigHookFunction( MIL_ID DigitizerID,
                                                 long HookType,
                                                 MDIGHOOKFCTPTR HookHandlerPtr,
                                                 void MPTYPE *UserDataPtr,
                                                 void MPTYPE *ThreadContextPtr);

MFTYPE32 PAPPHOOKFCTPTR MFTYPE PappHookFunction( long HookType,
                                                 PAPPHOOKFCTPTR HookHandlerPtr,
                                                 void MPTYPE *UserDataPtr,
                                                 void MPTYPE *ThreadContextPtr);

MFTYPE32 long MFTYPE PdispAlloc(        MIL_ID SystemID, 
                                        long DeviceNum, 
                                        char MPTYPE *DispFormat, 
                                        long InitFlag, 
                                        MIL_ID MPTYPE *IdVarPtr,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PdispSetLut(       MIL_ID DisplayID, 
                                        long   LutType, 
                                        long   Start,
                                        long   NbElem, 
                                        void  *LutData,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PdispOverlayKey(   MIL_ID DisplayID, 
                                        long KeyMode, 
                                        long KeyCond, 
                                        long KeyMask,
                                        long KeyColor,
                                        void MPTYPE *ThreadContextPtr);
                                        
MFTYPE32 void MFTYPE PdispControl(       MIL_ID DisplayID,
                                        long ControlType, 
                                        double ControlValue,
                                        void MPTYPE *ThreadContextPtr);
                                        
MFTYPE32 long MFTYPE PdispInquire(      MIL_ID DisplayID, 
                                        long InquireType, 
                                        void MPTYPE *UserVarPtr,
                                        void MPTYPE *ThreadContextPtr);
                                        
MFTYPE32 void MFTYPE PdispFree(         MIL_ID DisplayID,
                                        void MPTYPE *ThreadContextPtr);
                                                 
#ifndef MDISPHOOKFCTPTR

typedef MFTYPE32 long (MFTYPE MPTYPE *MDISPHOOKFCTPTR)(long HookType,
                                                       MIL_ID EventId,
                                                       void MPTYPE *UserPtr);
#endif


MFTYPE32 MDISPHOOKFCTPTR MFTYPE PdispHookFunction(MIL_ID DisplayID,
                                                  long HookType,
                                                  MDISPHOOKFCTPTR HookHandlerPtr,
                                                  void MPTYPE *UserDataPtr,
                                                  void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PdispZoom(         MIL_ID DisplayID, 
                                        long   XFactor, 
                                        long   YFactor,
                                        void MPTYPE *ThreadContextPtr);

MFTYPE32 void MFTYPE PdispPan(          MIL_ID DisplayID, 
                                        long   XOffset, 
                                        long   YOffset,
                                        void MPTYPE *ThreadContextPtr);

long MINTFTYPE PExtGrabHandler(         long FHookType, 
                                        MIL_ID EventId, 
                                        void MPTYPE *FExpansionFlagPtr);

/*==============================================================================*/
/*                                                                              */
/* PULSAR LowLevel prototypes                                                   */
/*                                                                              */
/*==============================================================================*/
MFTYPE32 void MFTYPE PLLinit(                   MIL_ID SystemId,
                                                void *OverlappedStruct);

MFTYPE32 long MFTYPE PLLsysAlloc(               long SystemType,
                                                long SystemNum,
                                                long InitFlag,
                                                MIL_ID *SystemIdPtr, 
                                                void *OverlappedStruct);

MFTYPE32 long MFTYPE PLLsysFree(                MIL_ID SystemId, 
                                                void *OverlappedStruct);

MFTYPE32 long MFTYPE PLLsysLutSelectPalette(    MIL_ID SystemId,
                                                unsigned short palette, 
                                                void *OverlappedStruct);

MFTYPE32 long MFTYPE PLLsysLutSelectMode(       MIL_ID SystemId,
                                                unsigned short mode, 
                                                void *OverlappedStruct);

MFTYPE32 long MFTYPE PLLsysLutWrite(            MIL_ID SystemId,
                                                unsigned short palette,
                                                unsigned short lut,
                                                unsigned short mode,
                                                unsigned char MPTYPE *pBuffer,
                                                unsigned long length,
                                                unsigned long offset, 
                                                void *OverlappedStruct);
   
MFTYPE32 long MFTYPE PLLdispOverlayKey(         MIL_ID SystemId, 
                                                long lKeyMode, 
                                                long lKeyCond,
                                                long nKeyColorHigh, 
                                                long nKeyColorLow, 
                                                void *OverlappedStruct);
   
MFTYPE32 long MFTYPE PLLsysControl(             MIL_ID SystemId,  
                                                long ControlType, 
                                                long ControlValue, 
                                                void *OverlappedStruct);

MFTYPE32 long MFTYPE PLLsysConfigAccess (MIL_ID  SystemId, 
                                          long    VendorId, 
                                          long    DeviceId,
                                          long    DeviceNumber,
                                          long    OpperationFlag,
                                          long    OpperationType,
                                          long    Offset,
                                          long    Size,
                                          void   *UserDataPtr,
                                          void   *OverlappedStruct);

MFTYPE32 long MFTYPE PLLdigAlloc (              MIL_ID SystemId, 
                                                char MPTYPE *DataFormat, 
                                                long InitFlag, 
                                                void *OverlappedStruct);
 
MFTYPE32 long MFTYPE PLLGetDcfFromDisk (        MIL_ID SystemId, 
                                                char MPTYPE *DataFormat, 
                                                long InitFlag, 
                                                void *OverlappedStruct);

MFTYPE32 long MFTYPE PLLSetDcfToHardWare (      MIL_ID SystemId, 
                                                long InitFlag, 
                                                void *OverlappedStruct);


MFTYPE32 void MFTYPE PLLdigFree(                MIL_ID SystemId,
                                                void *OverlappedStruct);

MFTYPE32 long MFTYPE PLLdigControl(             MIL_ID SystemId, 
                                                long ControlType, 
                                                double ControlValue,  
                                                void *OverlappedStruct);

MFTYPE32 long MFTYPE PLLdigChannel(             MIL_ID SystemId, 
                                                long Channel, 
                                                void *OverlappedStruct);    

MFTYPE32 long MFTYPE PLLdigReference(           MIL_ID SystemId, 
                                                long   ReferenceType, 
                                                long   ReferenceLevel, 
                                                void *OverlappedStruct);

MFTYPE32 void MFTYPE PLLdigGrabContinuous(      MIL_ID SystemId, 
                                                void *OverlappedStruct);   

MFTYPE32 long MFTYPE PLLdigHookFunction(        MIL_ID SystemId,
                                                long HookType,
                                                MDIGHOOKFCTPTR HookFunctionPtr,
                                                void MPTYPE *UserDataPtr, 
                                                void *OverlappedStruct);
                                                
MFTYPE32 long MFTYPE PLLdigGrabPosition(        MIL_ID SystemId,   
                                                long   SrcOffX,   
                                                long   SrcOffY,   
                                                long   SrcSizeX,   
                                                long   SrcSizeY,  
                                                long   DestSizeX,   
                                                long   DestSizeY,   
                                                long   DestPhysAddr, 
                                                long   DestPitchX,  
                                                long   DestPhysAddrU,   
                                                long   DestPitchXU,   
                                                long   DestPhysAddrV,  
                                                long   DestPitchXV, 
                                                long   GrabFlag,
                                                long  *StrideEven,
                                                long  *StrideOdd,
                                                void  *OverlappedStruct);
                                                
MFTYPE32 void MFTYPE PLLdigGrabSingle(          MIL_ID SystemId, 
                                                void  *OverlappedStruct);       

MFTYPE32 void MFTYPE PLLdigGrabStop(            MIL_ID SystemId,
                                                long   FieldToStop, 
                                                void  *OverlappedStruct);

MFTYPE32 long MFTYPE PLLsysInquire(             MIL_ID SystemId,
                                                long InquireType, 
                                                void  *OverlappedStruct);

MFTYPE32 long MFTYPE PLLdigInquire(             MIL_ID SystemId,
                                                long   InquireType, 
                                                void  *pInquireValue,
                                                void  *OverlappedStruct);

MFTYPE32 void MFTYPE PLLdigGrabField (          MIL_ID SystemId,
                                                long   FieldToGrab, 
                                                void  *OverlappedStruct);       

MFTYPE32 long MFTYPE PLLdigReadField (          MIL_ID SystemId, 
                                                void  *OverlappedStruct);

MFTYPE32 long MFTYPE PLLdigGrabStatus(          MIL_ID SystemId, 
                                                long StatusFlag, 
                                                void *OverlappedStruct); 

MFTYPE32 long MFTYPE PLLdigGrab (               MIL_ID SystemId,       
                                                long   WriteFormat,    
                                                long   SubsamplingX,   
                                                long   SubsamplingY,   
                                                long   TriggerMode,    
                                                long   OffsetX,        
                                                long   OffsetY,        
                                                long   InputSizeX,     
                                                long   InputSizeY,     
                                                long   DestSizeX,      
                                                long   DestSizeY,      
                                                long   BufferAdrs,     
                                                long   Pitch,          
                                                long   BufferAdrsU,    
                                                long   PitchU,         
                                                long   BufferAdrsV,    
                                                long   PitchV,         
                                                long   LutMode,        
                                                long   LutPalette,     
                                                long   GrabMode,
                                                long   GrabType,       
                                                long   GrabField,      
                                                long   StartOnField,   
                                                long   RouterValue,
                                                long   GrabStartMode,     
                                                long  *OddEvenFlagInverted,   
                                                long  *GrabInProgress, 
                                                void  *OverlappedStruct);

MFTYPE32 long MFTYPE PLLdigSetStartAdrs(        MIL_ID SystemId,       
                                                long   BufferAdrsEven,     
                                                long   BufferAdrsOdd,
                                                long   BufferAdrsEnd,
                                                long   BufferStrideEven,
                                                long   BufferStrideOdd,
                                                void  *OverlappedStruct);

MFTYPE32 short MFTYPE PLLDmaInit(void);
MFTYPE32 short MFTYPE PLLDmaClose(void);

MFTYPE32 long MFTYPE PLLDmaBufferAlloc(         MIL_ID        SystemId,
                                                unsigned long xsize, 
                                                unsigned long ysize,
                                                unsigned char pixelsize, 
                                                unsigned long Alignment,
                                                unsigned long Prefetch,
                                                unsigned long Frontier,
                                                unsigned long AllocFlags,
                                                unsigned long *AllocFlagsOk,
                                                void          *OverlappedStruct);
                                        
MFTYPE32 long MFTYPE PLLDmaBufferCreateId(      MIL_ID        SystemId, 
                                                unsigned long physaddr, 
                                                unsigned long xsize,
                                                unsigned long ysize, 
                                                unsigned char pixelsize, 
                                                unsigned long Alignment,
                                                unsigned long Prefetch,
                                                unsigned long Frontier,
                                                unsigned long AllocFlags,
                                                unsigned long *AllocFlagsOk,
                                                void          *OverlappedStruct);
                                            
MFTYPE32 void  MPTYPE * MFTYPE PLLDmaBufferGetPtr(MIL_ID SystemId, 
                                                  long   bufid, 
                                                  void  *OverlappedStruct);
                                                  
MFTYPE32 short MFTYPE PLLDmaBufferFree(         MIL_ID SystemId, 
                                                long   bufid, 
                                                void  *OverlappedStruct);

MFTYPE32 short MFTYPE  PLLDmaBufferLock(        long bufid);

MFTYPE32 short MFTYPE  PLLDmaBufferUnlock(      long bufid);

MFTYPE32 unsigned long MFTYPE PLLDmaBufferGetPhysAdr(MIL_ID SystemId, 
                                                     long   bufid, 
                                                     void  *OverlappedStruct);
                                                     
MFTYPE32 short MFTYPE  PLLDmaBufferFreeId(      MIL_ID SystemId, 
                                                long   bufid, 
                                                void  *OverlappedStruct);
                                             
MFTYPE32 long MFTYPE  PLLDmaBufferAssignId(     MIL_ID                SystemId,
                                                unsigned char MPTYPE *virtptr,
                                                unsigned long         physaddr, 
                                                unsigned long         xsize,
                                                unsigned long         ysize, 
                                                unsigned char         pixelsize, 
                                                unsigned long         Alignment,
                                                unsigned long         Prefetch,
                                                unsigned long         Frontier,
                                                unsigned long         AllocFlags,
                                                unsigned long         *AllocFlagsOk,
                                                void                 *OverlappedStruct);


MFTYPE32 long MFTYPE PLLdispAlloc (             MIL_ID       SystemId,
                                                char MPTYPE *DataFormat,
                                                long         InitFlag, 
                                                void        *OverlappedStruct);

MFTYPE32 long MFTYPE PLLdispControl(            MIL_ID       SystemId, 
                                                long         ControlType, 
                                                double       ControlValue, 
                                                void        *OverlappedStruct);

MFTYPE32 long MFTYPE PLLdispInquire(            MIL_ID       SystemId, 
                                                long         InquireType, 
                                                void        *OverlappedStruct);

MFTYPE32 long MFTYPE PLLdispSetLut (            MIL_ID                SystemId,
                                                unsigned long         LutType, 
                                                unsigned long MPTYPE *LutData, 
                                                void                 *OverlappedStruct);
                                                
MFTYPE32 long MFTYPE PLLdispZoom(               MIL_ID         SystemId, 
                                                unsigned long  XFactor, 
                                                unsigned long  YFactor, 
                                                void          *OverlappedStruct);
                                              
MFTYPE32 long MFTYPE PLLdispPan(                MIL_ID         SystemId, 
                                                unsigned long  XOffset, 
                                                unsigned long  YOffset, 
                                                void          *OverlappedStruct);

//==============================================================================
//
// internal prototype
//
//==============================================================================


MFTYPE32 unsigned long MFTYPE PLLFindPCIDevice( MIL_ID                SystemId,
                                                unsigned long         DeviceID, 
                                                unsigned long         VendorID, 
                                                unsigned long         Index, 
                                                unsigned long MPTYPE *PciId, 
                                                void                 *OverlappedStruct);

MFTYPE32 void * MFTYPE PLLFindPulsarStruc(      MIL_ID SystemId, 
                                                void  *OverlappedStruct);
   
MFTYPE32 long MFTYPE PLLFindPulsarId(           long SystemNum);

MFTYPE32 long MFTYPE PLLMapPhysicalMemory(      unsigned long                PhysAddr,
                                                unsigned long                Limit,
                                                void PULSARFARTYPE * MPTYPE *PointerToMemory,
                                                void MPTYPE                 *Handle);
   
MFTYPE32 long MFTYPE PLLUnmapPhysicalMemory(    void MPTYPE * Handle);

MFTYPE32 void MFTYPE PLLWriteConfigurationWord( unsigned long  BusDevFunNum, 
                                                unsigned long  RegNum,
                                                unsigned short Data);
                                                
MFTYPE32 unsigned short MFTYPE PLLReadConfigurationWord(MIL_ID        SystemID,
                                                        unsigned long BusDevFunNum,                                                         
                                                        unsigned long RegNum,
                                                        void         *OverlappedStruct);
                                                        
MFTYPE32 unsigned long MFTYPE PLLsysWait(       unsigned long value);

MFTYPE32 unsigned long MFTYPE PLLregReadMga_32( MIL_ID        SystemId,
                                                unsigned long offset,
                                                short         flag, 
                                                void         *OverlappedStruct);

MFTYPE32 void MFTYPE PLLregWriteMga_32(         MIL_ID         SystemId,
                                                unsigned long  offset, 
                                                unsigned long  data, 
                                                void          *OverlappedStruct);
   
MFTYPE32 unsigned short MFTYPE PLLregReadMga_16(MIL_ID        SystemId,
                                                unsigned long offset,
                                                short         flag, 
                                                void         *OverlappedStruct);

MFTYPE32 void MFTYPE PLLregWriteMga_16(         MIL_ID         SystemId,
                                                unsigned long  offset, 
                                                unsigned short data, 
                                                void          *OverlappedStruct);
   
MFTYPE32 unsigned char MFTYPE PLLregReadMga_8(  MIL_ID         SystemId,
                                                unsigned long  offset,
                                                short          flag, 
                                                void          *OverlappedStruct);

MFTYPE32 void MFTYPE PLLregWriteMga_8(          MIL_ID         SystemId,
                                                unsigned long  offset, 
                                                unsigned char  data, 
                                                void          *OverlappedStruct);
   
MFTYPE32 unsigned long MFTYPE PLLregRead7116_32(MIL_ID         SystemId,
                                                unsigned long  offset,
                                                short          flag, 
                                                void          *OverlappedStruct);

MFTYPE32 void MFTYPE PLLregWrite7116_32(        MIL_ID         SystemId,
                                                unsigned long  offset, 
                                                unsigned long  data, 
                                                void          *OverlappedStruct);
   
MFTYPE32 unsigned short MFTYPE PLLregRead7116_16(MIL_ID        SystemId,
                                                unsigned long  offset,
                                                short          flag, 
                                                void          *OverlappedStruct);

MFTYPE32 void MFTYPE PLLregWrite7116_16(        MIL_ID         SystemId,
                                                unsigned long  offset, 
                                                unsigned short data, 
                                                void          *OverlappedStruct);
   
MFTYPE32 unsigned char MFTYPE PLLregRead7116_8( MIL_ID         SystemId,
                                                unsigned long  offset,
                                                short          flag, 
                                                void          *OverlappedStruct);
   
MFTYPE32 void MFTYPE PLLregWrite7116_8(         MIL_ID         SystemId,
                                                unsigned long  offset, 
                                                unsigned char  data, 
                                                void          *OverlappedStruct);

/* C++ directive if needed */
#ifdef __cplusplus
}
#endif
#endif
