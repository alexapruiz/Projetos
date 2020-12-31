/*
    filename:   treg.h

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved

*/

#ifndef __TREG_H__
#define __TREG_H__

 
typedef struct{
   unsigned char buffer[4096];
   } MLB;

typedef struct hra
{

/* MAGIC memory mapped register list
   Generated on   October 27, 1992
*/

long  u000[16];            /* ------------- */

unsigned short c54cnt0;    /* M  : 040      */
unsigned short high_c54cnt0;
unsigned short c54cnt1;    /* M  : 044      */
unsigned short high_c54cnt1;
unsigned short c54cnt2;    /* M  : 048      */
unsigned short high_c54cnt2;
unsigned short c54ctrl;    /* M  : 04c      */
unsigned short high_c54ctrl;

long  u001[12];            /* ------------- */

unsigned short bt812wadd;  /* M  : 080      */
unsigned short high_bt812wadd;
unsigned short bt812res ;  /* M  : 084      */
unsigned short high_bt812res ;
unsigned short bt812data;  /* M  : 088      */
unsigned short high_bt812data;
unsigned short bt812radd;  /* M  : 08c      */
unsigned short high_bt812radd;

long  u002[14];            /* ------------- */

unsigned short bt261addr;  /* M  : 0c8      */
unsigned short high_bt261addr;
unsigned short bt261data;  /* M  : 0cc      */
unsigned short high_bt261data;

long  u003[12];            /* ------------- */

unsigned short bt473law ;  /* M  : 100      */
unsigned short high_bt473law ;
unsigned short bt473ldat;  /* M  : 104      */
unsigned short high_bt473ldat;
unsigned short bt473mask;  /* M  : 108      */
unsigned short high_bt473mask;
unsigned short bt473lar ;  /* M  : 10c      */
unsigned short high_bt473lar ;
unsigned short bt473oaw ;  /* M  : 110      */
unsigned short high_bt473oaw ;
unsigned short bt473odat;  /* M  : 114      */
unsigned short high_bt473odat;
unsigned short bt473vcmd;  /* M  : 118      */
unsigned short high_bt473vcmd;
unsigned short bt473oar ;  /* M  : 11c      */
unsigned short high_bt473oar ;

long  u004[8];             /* ------------- */

unsigned short dmaxstart;  /* M  : 140      */
unsigned short high_dmaxstart;
unsigned short dmaxstop ;  /* M  : 144      */
unsigned short high_dmaxstop ;
unsigned short dmaystart;  /* M  : 148      */
unsigned short high_dmaystart;
unsigned short dmactrl  ;  /* M  : 14c      */
unsigned short high_dmactrl  ;
unsigned short fbod     ;  /* M  : 150      */
unsigned short high_fbod     ;

long  u005[11];            /* ------------- */

unsigned short bt473cmd;   /* M  : 180      */
unsigned short high_bt473cmd;

long  u006[31];            /* ------------- */

unsigned short dclkctrl ;  /* M  : 200      */
unsigned short high_dclkctrl ;
unsigned short dpixfctrl;  /* M  : 204      */
unsigned short high_dpixfctrl;
unsigned short dsynctrl0;  /* M  : 208      */
unsigned short high_dsynctrl0;
unsigned short dsynctrl1;  /* M  : 20c      */
unsigned short high_dsynctrl1;
unsigned short dsintctrl;  /* M  : 210      */
unsigned short high_dsintctrl;
unsigned short dtimctrl ;  /* M  : 214      */
unsigned short high_dtimctrl ;
unsigned short duserctrl;  /* M  : 218      */
unsigned short high_duserctrl;
unsigned short dgainctrl;  /* M  : 21c      */
unsigned short high_dgainctrl;
unsigned short dpixpctrl;  /* M  : 220      */
unsigned short high_dpixpctrl;
unsigned short dgrabctrl;  /* M  : 224      */
unsigned short high_dgrabctrl;
unsigned short dstatus  ;  /* M  : 228      */
unsigned short high_dstatus  ;
unsigned short dlutctrl ;  /* M  : 22c      */
unsigned short high_dlutctrl ;

long  u007[2];             /* ------------- */

unsigned short fbor     ;  /* M  : 238      */
unsigned short high_fbor     ;
unsigned short vdoctrlr ;  /* M  : 23c      */
unsigned short high_vdoctrlr ;
unsigned short mlb1     ;  /* M  : 240      */
unsigned short high_mlb1     ;
unsigned short mlb2     ;  /* M  : 244      */
unsigned short high_mlb2     ;
unsigned short fboh     ;  /* M  : 248      */
unsigned short high_fboh     ;
unsigned short irqmask  ;  /* M  : 24c      */
unsigned short high_irqmask  ;
unsigned short irqclear ;  /* M  : 250      */
unsigned short high_irqclear ;
unsigned short irqstatus;  /* M  : 254      */
unsigned short high_irqstatus;

long  u008[10];            /* ------------- */

unsigned short keymask  ;  /* M  : 280      */
unsigned short high_keymask  ;
unsigned short keycolor ;  /* M  : 284      */
unsigned short high_keycolor ;


long u009[2];              /*-------------- */

unsigned short srtstarth;  /* M  : 290      */
unsigned short high_srtstarth;
unsigned short srtstartl;  /* M  : 294      */
unsigned short high_srtstartl;
unsigned short srtstartm;  /* M  : 298      */
unsigned short high_srtstartm;
unsigned short vdoctrl  ;  /* M  : 29c      */
unsigned short high_vdoctrl  ;
unsigned short bt254cmd ;  /* M  : 2a0      */
unsigned short high_bt254cmd ;
unsigned short bt254rref;  /* M  : 2a4      */
unsigned short high_bt254rref;
unsigned short bt254gref;  /* M  : 2a8      */
unsigned short high_bt254gref;
unsigned short bt254bref;  /* M  : 2ac      */
unsigned short high_bt254bref;
unsigned short bt254rlvl;  /* M  : 2b0      */
unsigned short high_bt254rlvl;
unsigned short bt254glvl;  /* M  : 2b4      */
unsigned short high_bt254glvl;
unsigned short bt254blvl;  /* M  : 2b8      */
unsigned short high_bt254blvl;
unsigned short bt254res ;  /* M  : 2bc      */
unsigned short high_bt254res ;
unsigned short dxstart  ;  /* M  : 2c0      */
unsigned short high_dxstart  ;
unsigned short dxstop   ;  /* M  : 2c4      */
unsigned short high_dxstop   ;
unsigned short dystart  ;  /* M  : 2c8      */
unsigned short high_dystart  ;
unsigned short dystop   ;  /* M  : 2cc      */
unsigned short high_dystop   ;
unsigned short fbo      ;  /* M  : 2d0      */
unsigned short high_fbo      ;

long  u010[11];            /* ------------- */

unsigned short mskdma   ;  /* M  : 300      */
unsigned short high_mskdma   ;
unsigned short rmskhost ;  /* M  : 304      */
unsigned short high_rmskhost ;
unsigned short wmskhost ;  /* M  : 308      */
unsigned short high_wmskhost ;
unsigned short wmskdig  ;  /* M  : 30c      */
unsigned short high_wmskdig  ;

long  u011[12];            /* ------------- */

unsigned short serctrl  ;  /* M  : 340      */
unsigned short high_serctrl  ;

long  u012[31];            /* ------------- */

unsigned short psgaddr  ;  /* M  : 3c0      */
unsigned short high_psgaddr  ;

long  u013[7];             /* ------------- */

unsigned short psgdata  ;  /* M  : 3e0      */
unsigned short high_psgdata  ;
} HRA;


typedef struct shadow
{
unsigned short dmactrl;
unsigned short dmaxstart;
unsigned short dmaxstop;
unsigned short dmaystart;

unsigned short dxstart;
unsigned short dxstop;
unsigned short dystart;
unsigned short dystop;

unsigned short keycolor;
unsigned short keymask;

unsigned short mskdma;
unsigned short rmskhost;
unsigned short wmskdig;
unsigned short wmskhost;
 
unsigned short srtstarth;
unsigned short srtstartm;
unsigned short srtstartl;

unsigned short dlutctrl;
unsigned short dpixfctrl;
unsigned short dsynctrl0;
} SHADOW;


/*  board memory mapping */

struct MagicBoard
{
unsigned char BdAvail;
HRA MFAR_1M * RegAddr;
MLB MFAR_1M * MlbAddr1;
MLB MFAR_1M * MlbAddr2;
unsigned long HighMapAdr;
void MHUGE * HighMapSel;
unsigned short DmaChannel;
};


struct MagicTsrDisp
{
unsigned long SrtStartl;
unsigned long SrtStartm;
unsigned long SrtStarth;
unsigned long DispVsyncInt;
unsigned long DispChanged;
};

#endif /* __TREG_H__ */
