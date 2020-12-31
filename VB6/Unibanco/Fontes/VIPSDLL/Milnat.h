/***************************************************************************/
/*

    Filename:  MILNAT.H
    Owner   :  Matrox Imaging dept.
    Rev     :  $Revision:   1.0  $
    Content :  This file contains the defines necessary to use the
               MIL NATIVE MODE DEVELOPER'S TOOLKIT.

    COPYRIGHT (c) 1993  Matrox Electronic Systems Ltd.
    All Rights Reserved

*/
/***************************************************************************/

/***************************************************************************/
/* MIL_USE_NATIVE_FLAG */

/* activate MIL native mode (if not done) */
#ifndef M_MIL_USE_NATIVE
#define M_MIL_USE_NATIVE   1
#endif

/* C++ directive if needed */
#ifdef __cplusplus
extern "C"
{
#endif

/***************************************************************************/
/* TOOLKIT FOR NATIVE MODE FUNCTION CALLS                                  */
/***************************************************************************/

MFTYPE32 long   MFTYPE MnatEnterNativeMode(MIL_ID SystemId, long ControlFlag);
MFTYPE32 long   MFTYPE MnatLeaveNativeMode(MIL_ID SystemId, long ControlFlag);
MFTYPE32 long   MFTYPE MnatModified(MIL_ID BufferId);
MFTYPE32 long   MFTYPE MnatModifiedRegion(MIL_ID BufferId, long OffX, long OffY, long SizeX, long SizeY);
MFTYPE32 long   MFTYPE MnatGetMilDriverId(MIL_ID Id);
MFTYPE32 MIL_ID MFTYPE MnatGetMilId(long OwnerSystemType,long OwnerSystemId,long ObjectAttribute,long ObjectDriverId);
MFTYPE32 long   MFTYPE MnatGetNativeId(MIL_ID Id);
MFTYPE32 long   MFTYPE MnatGetLocOrg(MIL_ID BufferId);
MFTYPE32 long   MFTYPE MnatGetLocSurf(MIL_ID BufferId);
MFTYPE32 long   MFTYPE MnatGetLocPosX(MIL_ID BufferId);
MFTYPE32 long   MFTYPE MnatGetLocPosY(MIL_ID BufferId);
MFTYPE32 long   MFTYPE MnatGetLocSizeX(MIL_ID BufferId);
MFTYPE32 long   MFTYPE MnatGetLocSizeY(MIL_ID BufferId);
MFTYPE32 long   MFTYPE MnatGetLocSizeBit(MIL_ID BufferId);
MFTYPE32 void   MPTYPE * MFTYPE MnatAccessSystemInfo(MIL_ID SystemId, long Operation,
                                                     long ModuleType,
                                                     void MPTYPE * MPTYPE *InfoPtr);


/***************************************************************************/
/* Some macro to replace function prototypes                               */
/***************************************************************************/
#define MfuncObjectAllocId     MfuncAllocId
#define MfuncObjectFreeId      MfuncFreeId
#define MfuncObjectGetType     MfuncIdGetObjectType
#define MfuncObjectGetUserPtr  MfuncIdGetUserPtr
#define MfuncObjectSetType     MfuncIdSetObjectType
#define MfuncObjectSetUserPtr  MfuncIdSetUserPtr

/***************************************************************************/
/* TOOLKIT FOR PSEUDO MIL FUNCTIONS DEVELOPMENT                            */
/***************************************************************************/

MFTYPE32 MIL_ID MFTYPE MfuncAlloc(char MPTYPE *FunctionName, long NumOfParam);
MFTYPE32 void MFTYPE MfuncParamId(MIL_ID FuncId,long ParamIndex,MIL_ID ParamVal,
                                  long ParamIs, long ParamHasAttr);
MFTYPE32 void MFTYPE MfuncParamChar(MIL_ID FuncId,long ParamIndex,char ParamVal);
MFTYPE32 void MFTYPE MfuncParamShort(MIL_ID FuncId,long ParamIndex,short ParamVal);
MFTYPE32 void MFTYPE MfuncParamLong(MIL_ID FuncId,long ParamIndex,long ParamVal);
MFTYPE32 void MFTYPE MfuncParamDouble(MIL_ID FuncId,long ParamIndex,double ParamVal);
MFTYPE32 void MFTYPE MfuncParamPointer(MIL_ID FuncId,long ParamIndex,void MPTYPE * ParamVal);
MFTYPE32 void MFTYPE MfuncParamString(MIL_ID FuncId,long ParamIndex,void MPTYPE * ParamVal);
MFTYPE32 long MFTYPE MfuncStart(MIL_ID FuncId);
MFTYPE32 long MFTYPE MfuncErrorReport(MIL_ID FuncId,
                                      long ErrorCode,
                                      char MPTYPE * ErrorString,
                                      char MPTYPE * ErrorSubString1,
                                      char MPTYPE * ErrorSubString2,
                                      char MPTYPE * ErrorSubString3);
MFTYPE32 long MFTYPE MfuncGetError(MIL_ID FunctId,long ErrorType,void MPTYPE *ErrorVarPtr);
MFTYPE32 void MFTYPE MfuncFreeAndEnd(MIL_ID FuncId);
MFTYPE32 long MFTYPE MfuncModified(MIL_ID ImageId);
MFTYPE32 long MFTYPE MfuncParamRegister(MIL_ID FuncId);
MFTYPE32 long MFTYPE MfuncParamCheck(MIL_ID FuncId);
MFTYPE32 MIL_ID MFTYPE MfuncAllocId(MIL_ID FunctionId, long ObjectType,
                                    void *UserPtr);
MFTYPE32 void MFTYPE MfuncFreeId(MIL_ID FunctionId, MIL_ID ObjectId);
MFTYPE32 void MFTYPE MfuncIdSetObjectType(MIL_ID FunctionId, MIL_ID ObjectId,
                                          long ObjectType);
MFTYPE32 long MFTYPE MfuncIdGetObjectType(MIL_ID FunctionId, MIL_ID ObjectId);
MFTYPE32 void MFTYPE MfuncIdSetUserPtr(MIL_ID FunctionId, MIL_ID ObjectId,
                                       void *UserPtr);
MFTYPE32 void* MFTYPE MfuncIdGetUserPtr(MIL_ID FunctionId, MIL_ID ObjectId);
MFTYPE32 long  MFTYPE MfuncPrintMessage(long Mode, char MPTYPE * Str1,
                                        char MPTYPE * Str2 ,
                                        char MPTYPE * Str3 ,
                                        char MPTYPE * Str4 ,
                                        char MPTYPE * Str5);
MFTYPE32 long  MFTYPE MfuncExit();
MFTYPE32 void MHUGETYPE * MFTYPE MfuncAllocMemory(long NbElements, long Size);
MFTYPE32 void             MFTYPE MfuncFreeMemory(void MHUGETYPE * MemoryP);


/***************************************************************************/
/* IMAGE SERIES SPECIFIC FUNCTION SET                                      */
/***************************************************************************/

MFTYPE32 MIL_ID MFTYPE MisDigToBufAlloc(MIL_ID DigitizerId,
                                        MIL_ID MPTYPE *BufferIdVarPtr);
MFTYPE32 void MFTYPE MisDigToBufFree(MIL_ID BufferId);
MFTYPE32 void MFTYPE MisBufCondImageAttach(MIL_ID BufferId,
                                           MIL_ID RoiId,
                                           long OffsetX,
                                           long OffsetY,
                                           long Condition,
                                           long CondValue);
MFTYPE32 void MFTYPE MisBufMask(MIL_ID BufferId, long Mask);
MFTYPE32 short MFTYPE MisNatGetLocGraph(MIL_ID BufferId, long Parameter);
MFTYPE32 void MFTYPE MisDigIteration(MIL_ID SystemId, long NbIteration);

/***************************************************************************/
/* HOST SPECIFIC FUNCTION SET                                              */
/***************************************************************************/

MFTYPE32 MIL_ID MFTYPE MnatBufCreate2d(MIL_ID SystemId,
                                       long SizeX,
                                       long SizeY,
                                       long Type,
                                       long Attribute,
                                       void MHUGETYPE *HostArrayPtr,
                                       MIL_ID MPTYPE *BufferIdVarPtr);
MFTYPE32 MIL_ID MFTYPE MnatBufCreateColor(MIL_ID SystemId,
                                          long SizeBand,
                                          long SizeX,
                                          long SizeY,
                                          long Type,
                                          long Attribute,
                                          void MHUGETYPE * MPTYPE * ArrayOfHostBufPtr,
                                          MIL_ID MPTYPE *BufferIdVarPtr);
MFTYPE32 void MFTYPE MnatBufDestroy(MIL_ID MilBufId);

MFTYPE32 long MFTYPE MnatBufGetPitch(MIL_ID BufferId);

MFTYPE32 void MHUGETYPE * MFTYPE MnatBufGetPointer(MIL_ID BufferId);


/***************************************************************************/
/* COMET SPECIFIC FUNCTION SET                                             */
/***************************************************************************/

MFTYPE32   void MFTYPE McometSysControl(MIL_ID SystemId,
                                        long ControlType,
                                        long ControlValue);

/* C++ directive if needed */
#ifdef __cplusplus
}
#endif

