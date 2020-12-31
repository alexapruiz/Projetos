/*
    filename:  proto.h
    Synopsis:
    Date:       March the 19th 1992

COPYRIGHT (c) 1990, 1991  Matrox Electronic Systems Ltd.
All Rights Reserved
*/

#ifndef __PROTO_H__
#define __PROTO_H__

#ifndef __TSERIES_H__
#include <tseries.h>
#endif /* __TSERIES_H__ */

//Some types defined here are used with prototypes
#ifndef __EXEC2_H__     //Don't include the standard library includes in
                        //MATROX internal test programs
#include <dos.h>
#include <stdio.h>
#endif

#include <tcamera.h>
#include <terror.h>

// Prototypes
#ifdef __cplusplus
extern "C" {
#endif

// prototypes

// get ptr functions
void PTYPE * FTYPE mGetPtr(unsigned short, unsigned short);

// init functions
int FTYPE mlBoardFPGALoad(int, const char PTYPE*);
int FTYPE mlInit(void);
int FTYPE mGetTSRVector(void);
int FTYPE mlInitDefault(void);
void FTYPE mlClose(void);
int FTYPE mlBoardSelect(int boardID);
struct MagicTsrBoard (MFAR_1M * FTYPE mlBoardGetConfig(void))[T_N_OF_BOARDS];


// prototypes for lut functions
void FTYPE mlSysLutGenerate (short ,unsigned char PTYPE *, short,long);
void FTYPE mlSysLutSet (short , short, short, short, unsigned char PTYPE *);
// void FTYPE mlSysLutSet (short , short, unsigned char PTYPE *);
void FTYPE mlSysLutGet (short, short, short, short, unsigned char PTYPE *);

// system inquire function
long FTYPE mlSysInquire(short info);

// memory organization
void FTYPE mlSysMemoryOrg(char org);
void FTYPE mlSysMasks(int resource, int value);

// IO functions
long FTYPE mlIOReadPixel(long surfaceID,long xposition,long yposition);
void FTYPE mlIOReadLine(long surfaceID,long xposition,long yposition,
           long direction,long npixel,void PTYPE * hostbuf);
void FTYPE mlIOReadArea(long surfaceID,long xstart,long ystart,
           long xend,long yend,void PTYPE * hostbuf);
void FTYPE mlIOReadHugeArea(long surfaceID,long xstart,long ystart,
           long xend,long yend,void MHUGE * hostbuf);
void FTYPE mlIOWritePixel(long surfaceID,long xposition,long yposition,
           long pixval);
void FTYPE mlIOWriteLine(long surfaceID,long xposition,long yposition,
           long direction,long npixel,void PTYPE * hostbuf);
void FTYPE mlIOWriteArea(long surfaceID,long xstart,long ystart,
           long xend,long yend,void PTYPE * hostbuf);
void FTYPE mlIOWriteRect(long surfaceID,long xstart,long ystart,
           long xend,long yend,long value);
void FTYPE mlIOWriteHugeArea(long surfaceID,long xstart,long ystart,
           long xend,long yend,void MHUGE * hostbuf);

void PTYPE * FTYPE mlIOGetLinePtr(long surfaceID,long xposition,
                   long yposition);
void FTYPE mlIOCopyArea(long srcSurf,long xStartSrc,long yStartSrc,
           long destSurf,long xStartDest,long yStartDest,long xSize,long ySize);

// display & keying
void FTYPE mlDispSelect (short dispMode,short xStart,short yStart);
void FTYPE mlDispKeying(short keyMode,char keyColor,char keyMask,
                        short keyCond);
short FTYPE mlDispVsyncEn(void);
short FTYPE mlDispVsyncDis(void);
void FTYPE mlDispVsyncWait(void);

// Camera functions
int FTYPE mlCamLoad(TCAMERA PTYPE* struc_reg, char PTYPE* filename);
int FTYPE  mlCamSave(char PTYPE* filename);
void FTYPE mlCamInit(void);
void FTYPE mlCamSelectChannel(unsigned short channel);
void FTYPE mlCamSelectSyncChannel(unsigned short channel);
void FTYPE mlCamGain(unsigned short gain);
void FTYPE mlCamClockDelay(unsigned short);
void FTYPE mlCamSelectPalette(unsigned short palette);
void FTYPE mlCamRefBlack(unsigned short level);
void FTYPE mlCamRefWhite(unsigned short ref);
void FTYPE mlCamFineGrnLvl(unsigned short level);
void FTYPE mlCamFineGrnRef(unsigned short ref);
short FTYPE mlCamUserBitOut(unsigned short userbitSelect, unsigned short value);
int FTYPE mlCamUserBitIn(void);
short FTYPE mlCamUserBitOutState(unsigned short state );
// This function is used by the MAGIC command interpretor only
int FTYPE mlCamSet(char PTYPE* filename);
void FTYPE mlCamSetHue (char hue);
void FTYPE mlCamSetContrast (unsigned char cont);
void FTYPE mlCamSetSatur (unsigned char sat);
void FTYPE mlCamSetBright (char level);
short FTYPE mlCamSetExpSig( unsigned short t0, unsigned short t1 );

// Grab control functions
void FTYPE mlGrabIteration(unsigned short condition,
                           unsigned short iterations);
void FTYPE mlGrabMode(unsigned short trig,
                      unsigned short pol);
void FTYPE mlGrabPosition(unsigned short surface,
                          unsigned short posX,
                          unsigned short posY);
void FTYPE mlGrabStart(void);
void FTYPE mlGrabStartSync(unsigned short activeBoard,unsigned short syncBoard);
void FTYPE mlGrabStop(unsigned short stopmode);
void FTYPE mlGrabStopSync(unsigned short stopmode,unsigned short activeBoard,
                          unsigned short syncBoard);
void FTYPE mlGrabSize(unsigned short x, unsigned short y);
void FTYPE mlGrabUnderSampling(unsigned short xsub, unsigned short ysub);
void FTYPE mlGrabVidPos(unsigned short xpos, unsigned short ypos);
void FTYPE mlCamWaitSync(void);
void FTYPE mlGrabInterlace(unsigned short storemode);

// functions to read write interrupt flags
unsigned long FTYPE mlReadInterruptCount(unsigned short intflag);
void FTYPE mlWriteInterruptCount(unsigned short intflag,unsigned long value);

// functions to disable enable interrupts
void FTYPE mlHsyncEnable(void);
void FTYPE mlHsyncDisable(void);

// internal functions
void FTYPE CallTsr(union REGS PTYPE *, struct SREGS PTYPE *);
void mSegToDsc ( void MFAR_1M* MFAR_1M* ptr );
void MHUGE * FTYPE mMapPhysicalMemory(unsigned long PhysicalAddress,
             unsigned long Size, unsigned long SegmentSize,
             unsigned long MFAR_1M* OldLinearAddress);
unsigned short mSetSelector(unsigned short Selector,
               unsigned long LinearAddress, unsigned long Size);
unsigned short FTYPE mFreeSelector(void MFAR * ptr);
unsigned short mFreePhysicalWat(unsigned long LinearAddress);
unsigned short mUnsetSelector(unsigned short Selector);
int mVerifyDPMIServer(void);
int FTYPE mHwInit(void);
void FTYPE mlCopyToDigitizer(TCAMERA PTYPE* ptrCamHigh);

void mLogDPMIError(char PTYPE* fctname, unsigned long dpmifct,
     unsigned short dpmicode);
void FTYPE mLogError(char PTYPE* fctname, unsigned int error_code);
unsigned long FTYPE CalcICD2061(unsigned long, unsigned long, unsigned long);
void FTYPE SetICD2061(char MFAR_1M*, unsigned long);
void stosdw(long MFAR* dst,long value,unsigned short count);
void movsdw(long MFAR* src, long MFAR* dst, unsigned short count);

// setregs prototypes
void FTYPE mlUpdateBit(unsigned int bit, unsigned int value);
void FTYPE mlRestoreFBOH(unsigned int word, unsigned int PTYPE * FBOH);
void FTYPE mlRestoreFBOD(unsigned int word, unsigned int PTYPE * FBOD);
void FTYPE mlRestoreFBO(unsigned int word, unsigned int PTYPE * FBO);
unsigned int FTYPE mlGetBit(unsigned int word,int n);
unsigned int FTYPE mlSetBit(unsigned int word, int n, unsigned int v);
unsigned int FTYPE mlGetBits(unsigned int word, int n, int k);
unsigned int FTYPE mlSetBits(unsigned int word, int n, int k, unsigned int v);

// core functions
int FTYPE mlopen(char PTYPE * fh, int mode);
int FTYPE ml_close(int file_h);
int FTYPE mlwrite(int fh,char PTYPE * buf,unsigned int nbytes);
int FTYPE mlread(int fh,char PTYPE * buf, unsigned int nbytes);
int FTYPE mlwritehuge(int fh,unsigned char huge * buf, unsigned char PTYPE *workbuf, unsigned int nbytes);
int FTYPE mlreadhuge (int fh,unsigned char huge * buf, unsigned char PTYPE *workbuf, unsigned int nbytes);
long FTYPE mlseek(int fh,long offset,int origin);
int FTYPE mlstrlen(char PTYPE * str);
char PTYPE * FTYPE mlstrcpy(char PTYPE *strdest, char PTYPE *strsrc);

// tiff functions prototypes

short FTYPE mlBuImportTiff(int );
short FTYPE mlBuExportTiff(int );
short FTYPE mlIOLoadTiff(char PTYPE * ,long ,long ,long );
short FTYPE mlIOStoreTiff(char PTYPE * ,long , long, long ,long ,long, long);
void FTYPE mlCompactBuff(char PTYPE *, char PTYPE *,int);
void FTYPE mlUnpackBuff(char PTYPE *,char PTYPE *, int);
void FTYPE mlCompactBuffHuge(char MHUGE * ,char PTYPE * ,int );
void FTYPE mlUnpackBuffHuge(char PTYPE * ,char MHUGE * ,int );
void FTYPE mlCompactBuffCol2Mono(short ,char MHUGE * ,char PTYPE * ,int );
void FTYPE mlUnpackBuffMono2Col(short , char PTYPE * ,char MHUGE * ,int );
void FTYPE mlInitTiffStruct(long ,long, long);
long FTYPE mlIOGetTiffInfo(char PTYPE *,short);
short FTYPE mlIOLoadTiff2Buf(char PTYPE *,short, short, short, void MHUGE *);
short FTYPE mlIOStoreBuf2Tiff(char PTYPE *, short, short, short, short, short, short, void MHUGE *);

// load the VCF file function
short FTYPE mlDispVcfLoad(char PTYPE * vcfname);

void  FTYPE MaskedVsync(void);

#ifdef MSC_NT
int  FTYPE mlAllocateDevice( unsigned short boards );
void  FTYPE mlIsrSpec( ULONG time[4] );
#endif /* MSC_NT */

#ifdef __cplusplus
}
#endif

#endif /* __T_PROTO_H__ */
