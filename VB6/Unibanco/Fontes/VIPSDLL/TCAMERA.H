/*
    filename:   tcamera.h
    Synopsis:   holds the definition of the TCAMERA structure used to hold
                everything about camera control.
    Date:       JULY the 16th 1992

COPYRIGHT (c) 1992 Matrox Electronic Systems Ltd.
All Rights Reserved.
*/


/*

#include <tcamera.h>

TCAMERA toto;

mlCamLoad(&toto, "filename.dlg");


toto.bt812.hwregs.cr0     is the cr0 register value in structure for Bt812
toto.bt812.allregs        get a pointer to an arrary holding all bt812
                          registers

// toto.bt812.allregs=(long MFAR *)&toto.bt812.hwregs.cr0
// toto.bt812.allregs[0];

toto.bt261.hwregs.clamp_start_low
                          is the <CLAMP start low> register value
                          in structure for Bt261
                          
toto.bt261.allregs        get a pointer to an array holding all bt261
                          registers
/
toto.dtimctrl             is the dtimctrl board register value in structure

mlCamSet(&toto)



NOTE: When a TCamera struct is allocated, all register values are
      initialized to -1 (0xffffffff).  When registers are loaded from
      file, the proper value is set in the long.  Since the bit 31
      is used to download or not a registers, only modified values are
      downloaded.  The register downloader will reset the bit 31 to
      1, don't load. 


*/

#ifndef __TCAMERA_H__
#define __TCAMERA_H__

/* ------------------------------------------------------------------- */

#define N_BT812_REGS  28
#define N_BT261_REGS  30
#define N_BT254_REGS  8
#define N_PSG_REGS    20
#define N_PT_REGS     6
#define N_BOARD_REGS  11


/* ------------------------------------------------------------------- */

typedef struct bt812regs 
   {
   long cr0;
   long cr2;
   long cr3;
   long cr4;
   long cr5;
   long cr6;
   long cr7;
   long brightness;
   long contrast;
   long saturation;
   long hue;
   long hclock_low;
   long hclock_high;
   long hdelay_low;
   long hdelay_high;
   long hactive_low;
   long hactive_high;

   long vdelay_low;
   long vdelay_high;
   long vactive_low;
   long vactive_high;
   long scar_reg0;
   long scar_reg1;
   long scar_reg2;
   long agc_delay;
   long burst_delay;
   long spl_rat_conv_l;
   long spl_rat_conv_h;
   } BT812REGS;


typedef struct bt261regs 
   {
   long cr0;
   long cr1;
   long cr2;
   long cr3;
   long sample;
   long osc_count_low;
   long osc_count_high;
   long unsued_0;
   long hsync_start_low;
   long hsync_start_high;
   long hsync_stop_low;
   long hsync_stop_high;
   long clamp_start_low;
   long clamp_start_high;
   long clamp_stop_low;
   long clamp_stop_high;
   long zero_start_low;
   long zero_start_high;
   long zero_stop_low;
   long zero_stop_high;
   long field_gate_start_low;
   long field_gate_start_high;
   long field_gate_stop_low;
   long field_gate_stop_high;
   long noise_gate_start_low;
   long noise_gate_start_high;
   long noise_gate_stop_low;
   long noise_gate_stop_high;
   long hcount_start_low;
   long hcount_start_high;
   } BT261REGS;

typedef struct bt254regs 
   {
   long cmd;
   long rref;
   long gref;
   long bref;
   long rlvl;
   long glvl;
   long blvl;
   long unsued_0;
   } BT254REGS;

typedef struct pt82c54regs 
   {
   long cnt0;
   long cnt1;
   long cnt2;
   long ctrl0;
   long ctrl1;
   long ctrl2;
   } PT82C54REGS; 

typedef struct psgregs
   {
   long dpyctl;
   long polctl;
   long extctl;
   long vtotal;
   long vsblnk;
   long veblnk;
   long vesync;
   long htotal;
   long hsblnk;
   long heblnk;
   long hesync;
   long vssyncs;
   long sethcnt;
   long setvcnt;
   long hsclmp;
   long heclmp;
   long hsbrst;
   long hebrst;
   long vsbrst;
   long vebrst;
   } PSGREGS; 

/* ------------------------------------------------------------------- */

typedef struct _bt812
   {
   BT812REGS   hwregs;
   long MFAR *  allregs;
   long       gru;
   } BT812;

typedef struct _bt261
   {
   BT261REGS   hwregs;
   long MFAR *  allregs;
   long       gru;
   } BT261;

typedef struct _bt254
   {
   BT254REGS   hwregs;
   long MFAR *  allregs;
   long       gru;
   } BT254;

typedef struct _pt
   {
   PT82C54REGS hwregs;
   long MFAR *  allregs;
   long       gru;
   } PT;

typedef struct _psg
   {
   PSGREGS     hwregs;
   long MFAR *  allregs;
   long       gru;
   } PSG;

typedef struct _info
   {
   long xsize;
   long ysize;
   long pixelclk;      // Now stored in Hz
   } INFO;

/* ------------------------------------------------------------------- */

typedef struct _tcamera
   {
   BT812    bt812;
   BT261    bt261;
   BT254    bt254;
   PT       ptimer;
   PSG      psg;
   long dclkctrl;
   long dpixfctrl;
   long dsynctrl0;
   long dsynctrl1;
   long dsintctrl;
   long dtimctrl;
   long duserctrl;
   long dgainctrl;
   long dpixpctrl;
   long userclock;
   INFO     caminfo;
   }  TCAMERA;  



/* ------------------------------------------------------------------- */
/* WARNING: Following structures are for test internal use only (test  */
/*          purpose don't remove                                       */

typedef struct _bt473regs 
   {
   long cmd; 
   long lar;
   long law;
   long ldat;
   long mask;
   long oar;
   long oaw;
   long odat;
   long unused_0;
   } BT473REGS;

/* ------------------------------------------------------------------- */

typedef struct _bt473
   {
   BT473REGS   hwregs;
   long far *  allregs;
   long       gru;
   } BT473;

/* ------------------------------------------------------------------- */

typedef struct _tenhmagic
   {
   BT473    bt473;

   long dgrabctrl;
   long dxstart;
   long dxstop;
   long dystart;
   long dystop;
   long dmactrl;
   long dmaxstart;
   long dmaxstop;
   long dmaystart;
   long eisaconf;
   long eisactrl;
   long fbo;
   long goutport;
   long irqclear;
   long irqmask;
   long keycolor;
   long keymask;
   long mlb1;
   long mlb2;
   long mskdma;
   long rmskhost;
   long srtstarth;
   long srtstartm;
   long srtstartl;
   long vdoctrl;
   long vlutar;
   long vlutaw;
   long vlutdat;
   long vlutmask;
   long wmskdig;
   long wmskhost;

   }  TENHMAGIC;  

#endif

